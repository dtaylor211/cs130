setup: requirements.txt
	pip install -r requirements.txt

clean:
	rm -rf __pycache__

test-workbook: 
	pytest -q ./tests/test_workbook.py

test:
	pytest -q ./tests/test_workbook.py
	pytest -q ./tests/test_sheet.py