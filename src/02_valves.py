    class Valves(BaseModel):
        base_url: Optional[str] = Field(
            default=None,
            description="Override the base URL for download links. Auto-detected from X-Original-Host header or WEBUI_URL env var if unset.",
        )
        templates: Optional[str] = Field(default="{}", description="JSON map of template names to content strings.")
        cleanup_schedule: Optional[str] = Field(default="{}", description="JSON schedule for auto-cleanup.")
        language: Optional[str] = Field(default="en", description="Language for error messages: en, pt, es, fr, de.")
        pass

