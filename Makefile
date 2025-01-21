# start virtual environment
.PHONY: start
start:
	python3 -m venv venv
	./venv/bin/pip install -r requirements.txt

# update requirements list
.PHONY: vendor-update
vendor-update:
	pipreqs --force .

# install requirements
.PHONY: vendor
vendor:
	pip3 install -r requirements.txt

# run
.PHONY: run
run:
	./venv/bin/python3 main.py