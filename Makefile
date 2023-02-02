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

test-performance:
	python -m cProfile -o program.prof ./tests/test_performance.py
	snakeviz program.prof