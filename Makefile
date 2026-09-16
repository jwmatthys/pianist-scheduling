LESSONS     ?= lesson_information.xlsx
PIANISTS    ?= pianist_availability.xlsx
JURY_INFO   ?= jury_information.xlsx
ASSIGNMENTS ?= $(shell ls -t $(basename $(LESSONS))_20*.xlsx 2>/dev/null | head -n1)
MARKDOWN_INPUT  ?= lesson_final_fa26.xlsx
MARKDOWN_OUTPUT ?= $(basename $(MARKDOWN_INPUT)).md
SEED        ?= 42

PYTHON ?= python3

.PHONY: install testdata pianist jury markdown clean distclean

install:
	$(PYTHON) -m pip install -r requirements.txt

testdata:
	$(PYTHON) generate_test_data.py --seed $(SEED)

pianist:
	$(PYTHON) generate_pianist_schedule.py --lessons $(LESSONS) --pianists $(PIANISTS)

jury:
	$(PYTHON) generate_jury_schedule.py --lessons $(LESSONS) --pianists $(PIANISTS) \
		--assignments $(ASSIGNMENTS) --jury-info $(JURY_INFO)

markdown:
	$(PYTHON) generate_lesson_markdown.py --input $(MARKDOWN_INPUT) --output $(MARKDOWN_OUTPUT)

clean:
	rm -f $(basename $(LESSONS))_20*.xlsx jury_schedule_*.xlsx \
		$(MARKDOWN_OUTPUT) lesson_final_fa26.md lesson_pianists_fa26.md

distclean: clean
	rm -f lesson_information.xlsx pianist_availability.xlsx jury_information.xlsx
