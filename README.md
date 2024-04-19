# Installation (Linux)
1. Create venv (`python -m venv venv`)
2. Activate venv (`source venv/bin/activate`)
3. Install dependencies (`pip install -r requirements.txt`)

# Usage
1. put your `*.xslx` files next to `main.py` file (can be changed with --inp argument)
2. Run script using venv `./venv/bin/python main.py`
3. log will be saved to `payment.log` (can be changed with --log argument)
4. Subdirectories with files will be created next to `main.py` file (can be changed with --dst argument)

use CRON job or alternative to create shedule for this task

you can use `./venv/bin/python main.py --help` for additional info

