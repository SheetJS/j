TARGET=j.js
FMT=xls xml xlsx xlsm xlsb csv slk dif txt ods misc
.PHONY: init
init:
	bash init.sh

.PHONY: clean
clean:
	if [ -e test_files ]; then rm -f test_files/*__.x*; fi

.PHONY: test mocha
test mocha: test.js
	mocha -R spec -t 10000

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

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)
	jscs $(TARGET)

.PHONY: cov cov-spin
cov: misc/coverage.html
cov-spin:
	make cov & bash misc/spin.sh $$!

COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov > $@

.PHONY: coveralls coveralls-spin
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js

coveralls-spin:
	make coveralls & bash misc/spin.sh $$!
