LESSONS     ?= lesson_information.xlsx
PIANISTS    ?= pianist_availability.xlsx
JURY_INFO   ?= jury_information.xlsx
ASSIGNMENTS ?= $(shell ls -t $(basename $(LESSONS))_20*.xlsx 2>/dev/null | head -n1)
SEED        ?= 42

PYTHON ?= python3

.PHONY: install testdata pianist jury clean distclean

install:
	$(PYTHON) -m pip install -r requirements.txt

testdata:
	$(PYTHON) generate_test_data.py --seed $(SEED)

pianist:
	$(PYTHON) generate_pianist_schedule.py --lessons $(LESSONS) --pianists $(PIANISTS)

jury:
	$(PYTHON) generate_jury_schedule.py --lessons $(LESSONS) --pianists $(PIANISTS) \
		--assignments $(ASSIGNMENTS) --jury-info $(JURY_INFO)

clean:
	rm -f $(basename $(LESSONS))_20*.xlsx jury_schedule_*.xlsx

distclean: clean
	rm -f lesson_information.xlsx pianist_availability.xlsx jury_information.xlsx
