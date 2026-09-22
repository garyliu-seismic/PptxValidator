EXE ?= C:\project_new\app-livedoc-construction-engine\src\core\PptxValidatorToolNet8\bin\Release\net10.0\PptxValidatorNet8.exe
FILE ?=

.PHONY: help validate validate-all autofix autofix-in-place clean

help:
	@echo "Targets:"
	@echo "  make validate FILE=deck.pptx           - diagnose a file or directory"
	@echo "  make validate-all FILE=./decks         - diagnose a directory recursively, including clean files"
	@echo "  make autofix FILE=deck.pptx            - auto-fix structural issues -> deck.fixed.pptx"
	@echo "  make autofix-in-place FILE=deck.pptx   - auto-fix in place (keeps a .bak)"
	@echo "  make clean                             - remove __pycache__ and stray .fixed/.bak files"
	@echo ""
	@echo "Override the validator exe path with EXE=path\\to\\PptxValidatorNet8.exe"

validate:
	@if [ -z "$(FILE)" ]; then echo "Usage: make validate FILE=deck.pptx"; exit 1; fi
	python scripts/validate.py "$(FILE)" --exe "$(EXE)"

validate-all:
	@if [ -z "$(FILE)" ]; then echo "Usage: make validate-all FILE=./decks"; exit 1; fi
	python scripts/validate.py "$(FILE)" --recursive --all --exe "$(EXE)"

autofix:
	@if [ -z "$(FILE)" ]; then echo "Usage: make autofix FILE=deck.pptx"; exit 1; fi
	python scripts/autofix.py "$(FILE)"

autofix-in-place:
	@if [ -z "$(FILE)" ]; then echo "Usage: make autofix-in-place FILE=deck.pptx"; exit 1; fi
	python scripts/autofix.py "$(FILE)" --in-place

clean:
	find . -name "__pycache__" -type d -prune -exec rm -rf {} +
	find . \( -name "*.fixed.pptx" -o -name "*.bak" \) -exec rm -f {} +
