    async def save_template(self, name: str, content: str) -> str:
        """Save a document template for reuse."""
        templates = json.loads(self.valves.templates or "{}")
        templates[name] = content
        self.valves.templates = json.dumps(templates)
        return f"Template '{name}' saved."

    async def use_template(self, name: str, __user__=None, __request__=None, **kwargs) -> str:
        """Generate a document from a saved template, replacing {placeholders}."""
        templates = json.loads(self.valves.templates or "{}")
        if name not in templates:
            return f"Template '{name}' not found. Available: {', '.join(templates.keys())}"
        content = templates[name]
        for key, value in kwargs.items():
            content = content.replace(f"{{{key}}}", str(value))
        return await self.generate_document(content, name, __user__=__user__, __request__=__request__)

    async def list_templates(self) -> str:
        """List all saved templates."""
        templates = json.loads(self.valves.templates or "{}")
        if not templates:
            return "No templates saved."
        result = "Available templates:\n"
        for name in templates:
            preview = templates[name][:50].replace('\n', ' ')
            result += f"- {name}: {preview}...\n"
        return result

    # --- v3.2.0: Scheduled Cleanup ---
