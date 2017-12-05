SHELL=/bin/bash
LIB=j
FMT=xlsx xlsm xlsb ods xls xml csv slk dif txt dbf misc
REQS=
ADDONS=
AUXTARGETS=
CMDS=bin/j.njs
HTMLLINT=

TARGET=$(LIB).js
FLOWTARGET=$(LIB).js
FLOWTGTS=$(TARGET) $(AUXTARGETS) $(AUXSCPTS)
UGLIFYOPTS=--support-ie8 -m
CLOSURE=/usr/local/lib/node_modules/google-closure-compiler/compiler.jar

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

j.exe: bin/j.njs j.js
	tail -n+2 $< | sed 's#\.\./#./j#g' > nexe.js
	nexe -i nexe.js -o $@
	rm nexe.js

.PHONY: pkg
pkg: bin/j.njs j.js ## Build pkg standalone executable
	pkg $<

## Testing

.PHONY: test mocha
test mocha: test.js ## Run test suite
	mocha -R spec -t 20000

#*                      To run tests for one format, make test_<fmt>
#*                      To run the core test suite, make test_misc
TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: travis
travis: ## Run test suite with minimal output
	mocha -R dot -t 30000


## Code Checking

.PHONY: fullint
fullint: lint old-lint flow ## Run all checks

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run eslint checks
	@eslint --ext .js,.njs,.json,.html,.htm $(TARGET) $(AUXTARGETS) $(CMDS) $(HTMLLINT) package.json bower.json
	if [ -e $(CLOSURE) ]; then java -jar $(CLOSURE) $(REQS) $(FLOWTARGET) --jscomp_warning=reportUnknownTypes >/dev/null; fi

.PHONY: old-lint
old-lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json test.js
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET) $(AUXTARGETS) test.js
	if [ -e $(CLOSURE) ]; then java -jar $(CLOSURE) $(REQS) $(FLOWTARGET) --jscomp_warning=reportUnknownTypes >/dev/null; fi

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
	mocha --require blanket -R html-cov -t 20000 > $@

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	mocha --require blanket --reporter mocha-lcov-reporter -t 20000 | node ./node_modules/coveralls/bin/coveralls.js

MDLINT=README.md
.PHONY: mdlint
mdlint: $(MDLINT) ## Check markdown documents
	alex $^
	mdspell -a -n -x -r --en-us $^

.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
