"""
Thin wrapper around Azure OpenAI for chat completions and embeddings.

Every other module should call chat_complete() / embed() from here rather
than importing the OpenAI SDK directly - this keeps the provider swappable
and the retry/error-handling logic in one place.
"""
from openai import AzureOpenAI

import config

_client = None


def get_client():
    """Lazily create and cache the AzureOpenAI client."""
    global _client
    if _client is None:
        config.validate_azure_openai()
        _client = AzureOpenAI(
            azure_endpoint=config.AZURE_OPENAI_ENDPOINT,
            api_key=config.AZURE_OPENAI_API_KEY,
            api_version=config.AZURE_OPENAI_API_VERSION,
        )
    return _client


def chat_complete(prompt, system=None, max_tokens=2000, retries=2):
    """
    Single-turn chat completion against the Azure chat deployment.

    Returns the response text. Raises the last exception if every
    attempt fails - callers are responsible for falling back.
    """
    client = get_client()
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})

    last_error = None
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model=config.AZURE_OPENAI_CHAT_DEPLOYMENT,
                messages=messages,
                max_completion_tokens=max_tokens,
            )
            return response.choices[0].message.content
        except Exception as e:
            last_error = e
            print(f"[AI_CLIENT] chat_complete attempt {attempt + 1} failed: {e}")

    raise last_error


def embed(texts):
    """
    Embed a single string or a list of strings using the Azure embedding
    deployment. Returns a single vector for a string input, or a list of
    vectors for a list input.
    """
    client = get_client()
    single_input = isinstance(texts, str)
    inputs = [texts] if single_input else list(texts)

    response = client.embeddings.create(
        model=config.AZURE_OPENAI_EMBEDDING_DEPLOYMENT,
        input=inputs,
    )
    vectors = [item.embedding for item in response.data]
    return vectors[0] if single_input else vectors
