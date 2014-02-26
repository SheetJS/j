FMT=xls xml xlsx xlsm xlsb misc

.PHONY: test mocha
test mocha: test.js
	mocha -R spec

TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: init
init:
	bash init.sh

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)

