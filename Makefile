publish:
	npm publish --access public

check_ts:
	npx tsc --noEmit

test_coverage:
	npx jest --coverage

test_badge: test_coverage
	npx jest-coverage-badges

ci: lint check_ts
	make test_coverage;
	make build;

lint:
	npx eslint src;

lint_fix:
	npx eslint --fix index.ts test/;



build: check_ts
	rm -rf dist;
	npx tsc

.PHONY: test