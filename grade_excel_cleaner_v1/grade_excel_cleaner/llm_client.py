from __future__ import annotations

from openai import OpenAI


def call_openai_compatible(
    *,
    base_url: str,
    api_key: str,
    model: str,
    system_prompt: str,
    user_prompt: str,
    temperature: float = 0.0,
) -> str:
    if not api_key:
        raise ValueError("请填写 API Key。")
    if not base_url:
        raise ValueError("请填写 LLM Base URL。")
    if not model:
        raise ValueError("请填写 Model Name。")

    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
        response_format={"type": "json_object"},
    )
    return response.choices[0].message.content or ""
