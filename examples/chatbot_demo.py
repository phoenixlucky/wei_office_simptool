"""对接本地 Ollama 的对话示例。

前置条件:
1. 本地已安装并启动 Ollama（默认 http://localhost:11434）
2. 已拉取模型: ollama pull llama3.2
"""

from wei_data_shu.ai import ChatBot


def main() -> None:
    bot = ChatBot(
        api_url="http://localhost:11434/api/chat",
        model="llama3.2",
        messages_file="messages.toml",
        history_file="chat_history.toml",
    )
    reply = bot.send_message("用一句话介绍你自己", stream=False)
    print("\n回复:", reply)


if __name__ == "__main__":
    main()
