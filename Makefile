setup: requirements.txt
	pip install -r requirements.txt

clean:
	rm -rf __pycache__

test-evaluator:
	pytest -q ./tests/test_evaluator.py
	pytest -q ./tests/test_evaluator_invalid.py

test:
	pytest -q ./tests/test_workbook.py
	pytest -q ./tests/test_sheet.py
	pytest -q ./tests/test_utils.py
	pytest -q ./tests/test_evaluator.py
	pytest -q ./tests/test_evaluator_invalid.py

test-performance-reference-chain:
	python -m cProfile -o program.prof \
		./tests/performance/test_reference_chain.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-circular-chain:
	python -m cProfile -o program.prof \
		./tests/performance/test_circular_chain.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-rename-volume1:
	python -m cProfile -o program.prof \
		./tests/performance/test_rename_volume1.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-rename-volume2:
	python -m cProfile -o program.prof \
		./tests/performance/test_rename_volume2.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-rename-chain:
	python -m cProfile -o program.prof \
		./tests/performance/test_rename_chain.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

PYLINT = pylint --rcfile=./.pylintrc
PYLINTFLAGS = -rn

PYTHONFILES := $(wildcard *.py)

pylint: $(patsubst %.py,%.pylint,$(PYTHONFILES))

lint-test-%:
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_$*.py

lint-%:
	$(PYLINT) $(PYLINTFLAGS) ./sheets/$*.py

lint-all:
	$(PYLINT) $(PYLINTFLAGS) ./sheets

