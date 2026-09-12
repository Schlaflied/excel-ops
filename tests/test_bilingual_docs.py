from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"


def test_every_public_doc_has_a_chinese_mirror_and_reciprocal_links():
    english_docs = sorted(
        path for path in DOCS.glob("*.md") if not path.name.endswith(".zh-CN.md")
    )
    assert english_docs, "expected at least one public documentation page"

    for english in english_docs:
        chinese = english.with_name(f"{english.stem}.zh-CN.md")
        assert chinese.exists(), f"missing Chinese mirror for {english.name}"

        english_text = english.read_text(encoding="utf-8")
        chinese_text = chinese.read_text(encoding="utf-8")
        assert english_text.startswith(f"[中文]({chinese.name}) | English\n")
        assert chinese_text.startswith(f"中文 | [English]({english.name})\n")


def test_readmes_link_to_docs_in_their_own_language():
    english_readme = (ROOT / "README.md").read_text(encoding="utf-8")
    chinese_readme = (ROOT / "README.zh-CN.md").read_text(encoding="utf-8")

    english_docs = sorted(
        path for path in DOCS.glob("*.md") if not path.name.endswith(".zh-CN.md")
    )
    for english in english_docs:
        chinese_name = f"{english.stem}.zh-CN.md"
        assert f"docs/{english.name}" in english_readme
        assert f"docs/{chinese_name}" in chinese_readme
        assert f"docs/{chinese_name}" not in english_readme
        assert f"docs/{english.name}" not in chinese_readme
