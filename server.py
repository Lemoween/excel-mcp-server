"""Excel MCP Server.
...
"""
from __future__ import annotations

import asyncio
import logging
import os
import sys
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from excel_mcp_server.bundle_registry import REGISTRY
from excel_mcp_server.tools.document import DocumentTools

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "excel_mcp.log"),
            encoding="utf-8",
        ),
        logging.StreamHandler(sys.stderr),
    ],
)
logger = logging.getLogger("excel_mcp")


class ExcelMCPServer:
    def __init__(self):
        self.server = Server("excel-mcp")
        self.registry = REGISTRY

        # Tool modules
        self.document_tools = DocumentTools(self.registry)

        self._tool_modules = [
            self.document_tools,
            # 未来: self.string_tools, self.number_tools, ...
        ]

        # Router
        self._tool_router = {}
        for module in self._tool_modules:
            for tool_def in module.get_tool_definitions():
                self._tool_router[tool_def["name"]] = module

        self._setup_handlers()

    def _setup_handlers(self):
        @self.server.list_tools()
        async def handle_list_tools():
            tools = []
            for module in self._tool_modules:
                for tool_def in module.get_tool_definitions():
                    tools.append(Tool(
                        name=tool_def["name"],
                        description=tool_def["description"],
                        inputSchema=tool_def["inputSchema"],
                    ))
            logger.info("Listed %d tools", len(tools))
            return tools

        @self.server.call_tool()
        async def handle_call_tool(name, arguments):
            arguments = arguments or {}
            logger.info("Tool call: %s(%s)", name, arguments)
            try:
                module = self._tool_router.get(name)
                if module is None:
                    return [TextContent(type="text", text=f"Unknown tool: '{name}'...")]
                result = module.execute(name, arguments)
                logger.info("Tool result length: %d", len(result))
                return [TextContent(type="text", text=result)]
            except Exception as e:
                logger.error("Tool error: %s", e, exc_info=True)
                return [TextContent(type="text", text=f"Error: {e}")]

    async def run(self):
        logger.info("Starting Excel MCP Server...")
        logger.info("Registered %d tools across %d modules",
                    len(self._tool_router), len(self._tool_modules))
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(read_stream, write_stream,
                                  self.server.create_initialization_options())


def main():
    server = ExcelMCPServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()