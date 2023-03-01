PYLINT = pylint --rcfile=./.pylintrc
PYLINTFLAGS = -rn

PYTHONFILES := $(wildcard *.py)

setup: requirements.txt
	pip install -r requirements.txt

clean:
	rm -rf __pycache__

test-evaluator:
	pytest -q ./tests/test_evaluator.py
	pytest -q ./tests/test_evaluator_invalid.py
	pytest -q ./tests/test_function_handler.py

test:
	pytest -q ./tests/test_workbook_basic.py
	pytest -q ./tests/test_workbook.py
	pytest -q ./tests/test_sheet.py
	pytest -q ./tests/test_utils.py
	pytest -q ./tests/test_evaluator.py
	pytest -q ./tests/test_evaluator_invalid.py
	pytest -q ./tests/test_function_handler.py

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

test-performance-rename-sheet-bulk:
	python -m cProfile -o program.prof \
		./tests/performance/test_rename_sheet_bulk.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-copy-sheet-bulk:
	python -m cProfile -o program.prof \
		./tests/performance/test_copy_sheet_bulk.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-copy:
	python -m cProfile -o program.prof \
		./tests/performance/test_copy.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-move:
	python -m cProfile -o program.prof \
		./tests/performance/test_move.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-move-cells-bulk:
	python -m cProfile -o program.prof \
		./tests/performance/test_move_cells_bulk.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-copy-cells-bulk:
	python -m cProfile -o program.prof \
		./tests/performance/test_copy_cells_bulk.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-fib:
	python -m cProfile -o program.prof \
		./tests/performance/test_fibonacci.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

test-performance-load:
	python -m cProfile -o program.prof \
		./tests/performance/test_load_wb.py
	@read -p "Visualize Data? [y/N] " ans && ans=$${ans:-N} ; \
    if [ $${ans} = y ] || [ $${ans} = Y ]; then \
        snakeviz program.prof; \
    fi

pylint: $(patsubst %.py,%.pylint,$(PYTHONFILES))

lint-test-context:
	$(PYLINT) $(PYLINTFLAGS) ./tests/context.py

lint-test-all:
	$(PYLINT) $(PYLINTFLAGS) ./tests/context.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_workbook.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_workbook_basic.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_sheet.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_evaluator.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_evaluator_invalid.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_utils.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_function_handler.py

lint-test-performance-all:
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_circular_chain.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_reference_chain.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_rename_volume1.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_rename_volume2.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_rename_chain.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_move.py
	$(PYLINT) $(PYLINTFLAGS) ./tests/performance/test_copy.py

lint-test-%:
	$(PYLINT) $(PYLINTFLAGS) ./tests/test_$*.py

lint-%:
	$(PYLINT) $(PYLINTFLAGS) ./sheets/$*.py

lint-all:
	$(PYLINT) $(PYLINTFLAGS) ./sheets

