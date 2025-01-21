# start virtual environment
start:
	python3 -m venv venv
	./venv/bin/pip install -r requirements.txt

# update requirements list
vendor-update:
	pipreqs --force .

# install requirements
vendor:
	pip3 install -r requirements.txt

# run
run: start
	./venv/bin/python3 main.py