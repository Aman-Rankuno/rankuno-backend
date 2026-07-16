"""
AI provider adapters for the Chat with AI feature.

Each adapter takes the same provider-neutral inputs (system prompt,
message history, tool schemas, a tool-executor callback) and handles
that provider's specific tool-calling protocol and message format
internally. Adding a new provider means writing one new adapter class
here - nothing else in the chat endpoint needs to change.
"""
import json
from app.config import settings
from app.services.ai_chat.tools import TOOL_SCHEMAS, execute_tool

MAX_TOOL_ROUNDS = 5  # safety cap: stop even if a model keeps calling tools


class ProviderNotConfigured(Exception):
    pass


class OpenAIProvider:
    name = "openai"

    def __init__(self):
        if not settings.OPENAI_API_KEY:
            raise ProviderNotConfigured("OPENAI_API_KEY is not set.")
        from openai import OpenAI
        self.client = OpenAI(api_key=settings.OPENAI_API_KEY)

    def _tool_defs(self):
        return [
            {
                "type": "function",
                "function": {
                    "name": t["name"],
                    "description": t["description"],
                    "parameters": t["parameters"],
                },
            }
            for t in TOOL_SCHEMAS
        ]

    def chat(self, system_prompt: str, history: list, crawl) -> str:
        messages = [{"role": "system", "content": system_prompt}] + history

        for _ in range(MAX_TOOL_ROUNDS):
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                tools=self._tool_defs(),
            )
            choice = response.choices[0]
            msg = choice.message

            if not msg.tool_calls:
                return msg.content or ""

            messages.append({
                "role": "assistant",
                "content": msg.content,
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                    }
                    for tc in msg.tool_calls
                ],
            })

            for tc in msg.tool_calls:
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except json.JSONDecodeError:
                    args = {}
                result = execute_tool(tc.function.name, crawl, args)
                messages.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "content": json.dumps(result),
                })

        return "I wasn't able to finish answering that within the allowed number of tool calls. Try asking a more specific question."


class AnthropicProvider:
    name = "claude"

    def __init__(self):
        if not settings.ANTHROPIC_API_KEY:
            raise ProviderNotConfigured("ANTHROPIC_API_KEY is not set.")
        import anthropic
        self.client = anthropic.Anthropic(api_key=settings.ANTHROPIC_API_KEY)

    def _tool_defs(self):
        return [
            {
                "name": t["name"],
                "description": t["description"],
                "input_schema": t["parameters"],
            }
            for t in TOOL_SCHEMAS
        ]

    def chat(self, system_prompt: str, history: list, crawl) -> str:
        messages = list(history)

        for _ in range(MAX_TOOL_ROUNDS):
            response = self.client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1024,
                system=system_prompt,
                messages=messages,
                tools=self._tool_defs(),
            )

            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                text_blocks = [b.text for b in response.content if b.type == "text"]
                return "\n".join(text_blocks)

            messages.append({"role": "assistant", "content": response.content})

            tool_results = []
            for tu in tool_uses:
                result = execute_tool(tu.name, crawl, tu.input or {})
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": tu.id,
                    "content": json.dumps(result),
                })
            messages.append({"role": "user", "content": tool_results})

        return "I wasn't able to finish answering that within the allowed number of tool calls. Try asking a more specific question."


PROVIDERS = {
    "openai": OpenAIProvider,
    "claude": AnthropicProvider,
}


def get_provider(name: str):
    cls = PROVIDERS.get(name)
    if not cls:
        raise ValueError(f"Unknown provider: {name}")
    return cls()
