from pathlib import Path


ROOT = Path(__file__).parents[1]


def test_agent_skill_has_required_entrypoints():
    skill = ROOT / ".agents" / "skills" / "excel-agent" / "SKILL.md"
    metadata = skill.parent / "agents" / "openai.yaml"
    assert skill.exists()
    assert metadata.exists()
    text = skill.read_text(encoding="utf-8")
    assert "name: excel-agent" in text
    assert "references/shared.md" in text
    assert "references/workflows.md" in text
    assert "references/connectors.md" in text


def test_agent_references_exist():
    base = ROOT / ".agents" / "skills" / "excel-agent" / "references"
    for name in ("shared.md", "workflows.md", "connectors.md"):
        assert (base / name).exists()
