# OpenCode-Pi Bridge Execution Rules
- You are running inside Pi Coding Agent via the `opencode-pi` CLI bridge.
- CRITICAL: Never attempt to invoke native OpenCode tools (like `read`, `write`, `bash`, or `edit`) directly. Doing so will immediately crash the agent turn.
- If you need to read a file, look at the directory, or execute a command, you must strictly wrap the payload inside Pi's tool control syntax:
  <pi_tool_call>{"name": "TOOL_NAME", "arguments": { ... }}</pi_tool_call>