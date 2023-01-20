setup: requirements.txt
	pip install -r requirements.txt

clean:
	rm -rf __pycache__

test-workbook: 
	pytest -q ./tests/test_workbook.py

test-evaluator:
	pytest -q ./tests/test_evaluator.py

test-evaluator-invalid:
	pytest -q ./tests/test_evaluator_invalid.py

test:
	pytest -q ./tests/test_workbook.py
	pytest -q ./tests/test_sheet.py
	pytest -q ./tests/test_evaluator.py
	pytest -q ./tests/test_evaluator_invalid.py