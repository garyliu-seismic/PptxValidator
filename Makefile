EXE  ?= C:\project_new\app-livedoc-construction-engine\src\core\PptxValidatorToolNet8\bin\Release\net10.0\PptxValidatorNet8.exe
FILE ?=

.PHONY: help validate validate-all validate-dotnet autofix autofix-in-place test clean

help:
	@echo "Targets:"
	@echo "  make validate FILE=deck.pptx           - validate with pure-Python backend (default)"
	@echo "  make validate-all FILE=./decks         - validate directory recursively, show all files"
	@echo "  make validate-dotnet FILE=deck.pptx    - validate with .NET backend (requires EXE)"
	@echo "  make autofix FILE=deck.pptx            - auto-fix structural issues -> deck.fixed.pptx"
	@echo "  make autofix-in-place FILE=deck.pptx   - auto-fix in place (keeps a .bak)"
	@echo "  make test                              - run gap-comparison test vs .NET backend"
	@echo "  make clean                             - remove __pycache__ and stray .fixed/.bak files"
	@echo ""
	@echo "Override the .NET exe path with EXE=path\\to\\PptxValidatorNet8.exe"

validate:
	@if [ -z "$(FILE)" ]; then echo "Usage: make validate FILE=deck.pptx"; exit 1; fi
	python scripts/validate.py "$(FILE)"

validate-all:
	@if [ -z "$(FILE)" ]; then echo "Usage: make validate-all FILE=./decks"; exit 1; fi
	python scripts/validate.py "$(FILE)" --recursive --all

validate-dotnet:
	@if [ -z "$(FILE)" ]; then echo "Usage: make validate-dotnet FILE=deck.pptx"; exit 1; fi
	python scripts/validate.py "$(FILE)" --exe "$(EXE)"

autofix:
	@if [ -z "$(FILE)" ]; then echo "Usage: make autofix FILE=deck.pptx"; exit 1; fi
	python scripts/autofix.py "$(FILE)"

autofix-in-place:
	@if [ -z "$(FILE)" ]; then echo "Usage: make autofix-in-place FILE=deck.pptx"; exit 1; fi
	python scripts/autofix.py "$(FILE)" --in-place

test:
	python scripts/test_compare.py --exe "$(EXE)"

clean:
	find . -name "__pycache__" -type d -prune -exec rm -rf {} +
	find . \( -name "*.fixed.pptx" -o -name "*.bak" \) -exec rm -f {} +
