"""
Talent marketplace matching engine.

Maps a candidate's validated skill profile onto:
  - external job openings        (candidate-job matching)
  - internal mobility roles      (internal mobility)
  - mentors                      (mentorship pairing)
  - learning resources           (learning suggestions)
  - a suggested next-step career path

Scoring is a weighted blend, which keeps the result explainable (we can
always say WHY something matched) rather than an opaque single number:

  skill match      70%   - dominant signal; the skills are the job requirement
  assessment score 20%   - demonstrated proficiency, not just claimed skills
  experience        10%   - breadth proxy (how many distinct skills they have)

The skill-match component itself combines two views:
  - exact/token overlap    (precise, but blind to synonyms)
  - embedding similarity   (catches "JS" ~ "JavaScript", "ML" ~ "Machine
                            Learning") via text-embedding-3-large

Embeddings are computed lazily and cached in the DB, so the first match
for a new job/mentor/resource pays the embedding cost once and every
later match is a cheap vector comparison.
"""
import json
import math

import db
from services import ai_client

# Weights (must sum to 1.0)
W_SKILLS = 0.70
W_ASSESSMENT = 0.20
W_EXPERIENCE = 0.10

# Above this many distinct skills, the experience component is maxed out.
EXPERIENCE_SATURATION = 8


def _cosine_similarity(a, b):
    if not a or not b:
        return 0.0
    dot = sum(x * y for x, y in zip(a, b))
    norm_a = math.sqrt(sum(x * x for x in a))
    norm_b = math.sqrt(sum(y * y for y in b))
    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot / (norm_a * norm_b)


def _tokenize_skills(skills_text):
    """Normalize a comma-separated skill string into a lowercase set."""
    if not skills_text:
        return set()
    return {s.strip().lower() for s in skills_text.split(",") if s.strip()}


def _token_overlap(candidate_skills, required_skills):
    """Fraction of the ROLE's requirements the candidate covers. Deliberately
    asymmetric: having 20 unrelated extra skills shouldn't inflate the score
    for a role that needs 3 specific ones."""
    if not required_skills:
        return 0.0
    matched = 0
    for req in required_skills:
        # substring match in either direction catches "aws" vs "aws cloud"
        if any(req in cand or cand in req for cand in candidate_skills):
            matched += 1
    return matched / len(required_skills)


def _get_or_compute_embedding(row, text, setter):
    """Return the cached embedding for a row, computing and persisting it
    on first use. Returns None if the embedding service is unavailable -
    callers fall back to token overlap alone."""
    cached = row.get("embedding_json")
    if cached:
        try:
            return json.loads(cached)
        except (ValueError, TypeError):
            pass  # corrupt cache - recompute below

    try:
        vector = ai_client.embed(text)
        setter(row["id"], json.dumps(vector))
        return vector
    except Exception as e:
        print(f"[MATCHING] embedding unavailable for row {row.get('id')}: {e}")
        return None


def _embed_candidate_profile(skills_text):
    try:
        return ai_client.embed(skills_text)
    except Exception as e:
        print(f"[MATCHING] candidate embedding unavailable: {e}")
        return None


def _skill_match_score(candidate_skills, candidate_vector, required_skills_text, item_vector):
    """Blend exact token overlap with embedding similarity. If embeddings
    are unavailable, fall back to token overlap alone rather than failing."""
    overlap = _token_overlap(candidate_skills, _tokenize_skills(required_skills_text))

    if candidate_vector and item_vector:
        similarity = _cosine_similarity(candidate_vector, item_vector)
        # cosine on this model is ~0.2-0.9 for related text; rescale so the
        # usable band spreads across 0-1 instead of bunching up high
        similarity = max(0.0, min(1.0, (similarity - 0.2) / 0.6))
        return 0.5 * overlap + 0.5 * similarity

    return overlap


def _experience_score(skill_count):
    return min(1.0, skill_count / EXPERIENCE_SATURATION)


def build_candidate_profile(user_id):
    """Gather everything the matching engine needs about one candidate."""
    skill_rows = db.get_skills_for_user(user_id)
    skill_names = []
    seen = set()
    for row in skill_rows:
        name = (row["skill_name"] or "").strip()
        if name and name.lower() not in seen:
            seen.add(name.lower())
            skill_names.append(name)

    assessment = db.get_latest_completed_assessment(user_id)
    accuracy = float(assessment["accuracy"]) / 100.0 if assessment else 0.0

    return {
        "skills": skill_names,
        "skills_text": ", ".join(skill_names),
        "assessment_accuracy": accuracy,
        "has_assessment": assessment is not None,
    }


def match_jobs(profile, internal=False, top_n=5):
    """Rank job openings against a candidate profile.

    internal=False -> external job matching
    internal=True  -> internal mobility opportunities
    """
    candidate_skills = _tokenize_skills(profile["skills_text"])
    candidate_vector = _embed_candidate_profile(profile["skills_text"]) if profile["skills_text"] else None
    experience = _experience_score(len(profile["skills"]))

    results = []
    for job in db.get_jobs():
        if bool(job["is_internal"]) != internal:
            continue

        job_text = f"{job['title']}. {job['description']} Required skills: {job['required_skills']}"
        job_vector = _get_or_compute_embedding(job, job_text, db.set_job_embedding)

        skill_score = _skill_match_score(
            candidate_skills, candidate_vector, job["required_skills"], job_vector
        )
        total = (
            W_SKILLS * skill_score
            + W_ASSESSMENT * profile["assessment_accuracy"]
            + W_EXPERIENCE * experience
        )

        required = _tokenize_skills(job["required_skills"])
        matched_skills = sorted(
            {req for req in required if any(req in c or c in req for c in candidate_skills)}
        )
        missing_skills = sorted(required - set(matched_skills))

        results.append({
            "id": job["id"],
            "title": job["title"],
            "department": job["department"],
            "description": job["description"],
            "required_skills": job["required_skills"],
            "match_percent": round(total * 100, 1),
            "skill_score": round(skill_score * 100, 1),
            "matched_skills": matched_skills,
            "missing_skills": missing_skills,
        })

    results.sort(key=lambda r: r["match_percent"], reverse=True)
    return results[:top_n]


def match_mentors(profile, top_n=3):
    """Suggest mentors. Deliberately biased toward COMPLEMENTARY expertise -
    a mentor is most useful where the candidate is weakest, so we rank on
    overlap with the candidate's growth areas as well as their strengths."""
    candidate_skills = _tokenize_skills(profile["skills_text"])
    candidate_vector = _embed_candidate_profile(profile["skills_text"]) if profile["skills_text"] else None

    results = []
    for mentor in db.get_mentors():
        mentor_text = f"{mentor['name']}. Expertise: {mentor['expertise']}. {mentor['bio']}"
        mentor_vector = _get_or_compute_embedding(mentor, mentor_text, db.set_mentor_embedding)

        relevance = _skill_match_score(
            candidate_skills, candidate_vector, mentor["expertise"], mentor_vector
        )
        expertise = _tokenize_skills(mentor["expertise"])
        shared = sorted({e for e in expertise if any(e in c or c in e for c in candidate_skills)})
        can_teach = sorted(expertise - set(shared))

        results.append({
            "id": mentor["id"],
            "name": mentor["name"],
            "expertise": mentor["expertise"],
            "bio": mentor["bio"],
            "relevance_percent": round(relevance * 100, 1),
            "shared_skills": shared,
            "can_teach": can_teach,
        })

    results.sort(key=lambda r: r["relevance_percent"], reverse=True)
    return results[:top_n]


def suggest_learning(profile, missing_skills=None, top_n=5):
    """Recommend learning resources. Skills the candidate is MISSING for
    their best-matching roles are prioritized over ones they already have -
    the point is closing gaps, not reinforcing strengths."""
    candidate_skills = _tokenize_skills(profile["skills_text"])
    gap_skills = {s.strip().lower() for s in (missing_skills or []) if s.strip()}

    results = []
    for resource in db.get_learning_resources():
        tag = (resource["skill_tag"] or "").strip().lower()
        title = (resource["title"] or "").strip().lower()

        # Check the title as well as the tag: if the tag vocabulary ever
        # drifts from jobs.required_skills, a gap still resolves via the
        # title rather than silently scoring as "not a gap".
        def _mentions(skill):
            return skill in tag or tag in skill or skill in title

        is_gap = any(_mentions(g) for g in gap_skills) if gap_skills else False
        is_known = any(tag in c or c in tag for c in candidate_skills)

        if is_gap:
            priority, reason = 3, "Closes a gap for your top role matches"
        elif is_known:
            priority, reason = 2, "Deepens a skill you already have"
        else:
            priority, reason = 1, "Broadens your profile"

        results.append({
            "id": resource["id"],
            "title": resource["title"],
            "skill_tag": resource["skill_tag"],
            "url": resource["url"],
            "description": resource["description"],
            "priority": priority,
            "reason": reason,
        })

    results.sort(key=lambda r: r["priority"], reverse=True)
    return results[:top_n]


def suggest_career_path(profile, top_jobs):
    """Produce a short, concrete next-step narrative. Uses the AI when
    available (it can reason about progression), but always falls back to
    a deterministic summary built from the match data so this section is
    never empty."""
    if not top_jobs:
        return "Complete an assessment to receive career path recommendations."

    best = top_jobs[0]
    gaps = ", ".join(best["missing_skills"][:4]) or "no major gaps"

    prompt = f"""
A candidate has these skills: {profile['skills_text'] or 'none recorded'}
Their assessment accuracy was {profile['assessment_accuracy'] * 100:.0f}%.
Their strongest role match is "{best['title']}" ({best['match_percent']}% match).
Skills they are missing for it: {gaps}.

In 3-4 sentences, describe a realistic career path for this candidate:
their likely next role, what to strengthen first, and a plausible step
after that. Be specific and practical. Do not use bullet points.
"""
    try:
        text = ai_client.chat_complete(prompt, max_tokens=3000).strip()
        if text:
            return text
    except Exception as e:
        print(f"[MATCHING] career path AI unavailable: {e}")

    return (
        f"Your strongest match is {best['title']} ({best['match_percent']}% fit) in "
        f"{best['department']}. To strengthen this path, focus next on: {gaps}. "
        f"Building those areas would also open up adjacent roles in {best['department']}."
    )


def get_recommendations(user_id):
    """Full marketplace recommendation set for one candidate - this is the
    single entry point the UI calls."""
    profile = build_candidate_profile(user_id)

    external_jobs = match_jobs(profile, internal=False, top_n=5)
    internal_jobs = match_jobs(profile, internal=True, top_n=5)

    # Learning priorities are driven by the gaps in the best matches
    gap_pool = []
    for job in (external_jobs[:2] + internal_jobs[:2]):
        gap_pool.extend(job["missing_skills"])

    return {
        "profile": profile,
        "external_jobs": external_jobs,
        "internal_jobs": internal_jobs,
        "mentors": match_mentors(profile, top_n=3),
        "learning": suggest_learning(profile, missing_skills=gap_pool, top_n=5),
        "career_path": suggest_career_path(profile, external_jobs or internal_jobs),
    }
