SHELL=/bin/bash
LIB=j
FMT=xlsx xlsm xlsb ods xls xml csv slk dif txt dbf misc
REQS=
ADDONS=
AUXTARGETS=
CMDS=bin/j.njs
HTMLLINT=

TARGET=$(LIB).js
UGLIFYOPTS=--support-ie8

## Main Targets

.PHONY: all
all: init ## .

.PHONY: clean
clean: ## Remove targets and build artifacts
	rm -f tmp/*__.*

.PHONY: init
init: ## Initial setup for development
	bash init.sh

.PHONY: graph
graph: formats.png ## Rebuild format conversion graph
formats.png: formats.dot
	circo -Tpng -o$@ $<

.PHONY: nexe
nexe: j.exe ## Build nexe standalone executable

j.exe: bin/j.njs
	nexe -i $< -o $@ --flags

## Testing

.PHONY: test mocha
test mocha: test.js ## Run test suite
	mocha -R spec -t 30000

#*                      To run tests for one format, make test_<fmt>
TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: 2011
2011:
	./tests/open_excel_2011.sh

.PHONY: numbers
numbers:
	./tests/open_numbers.sh

## Code Checking

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET) $(AUXTARGETS)

.PHONY: flow
flow: lint ## Run flow checker
	@flow check --all --show-all-errors

.PHONY: cov
cov: misc/coverage.html ## Run coverage test

#*                      To run coverage tests for one format, make cov_<fmt>
COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov -t 30000 > $@

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	mocha --require blanket --reporter mocha-lcov-reporter -t 30000 | node ./node_modules/coveralls/bin/coveralls.js


.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
