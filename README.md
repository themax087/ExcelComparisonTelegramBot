# Excel Comparison Bot
This is a Telegram bot that helps you compare two Excel files and shows the differences between them.

The bot supports .xlsx and .xls file formats and highlights the differences in red. It's a quick and simple way to compare spreadsheets without having to manually inspect each cell.

## Usage
To start using the bot, send the /compare command and follow the instructions. You'll need to send two Excel files to the bot to compare. Once the comparison is done, the bot will send you a new Excel file with the differences highlighted in red.

### Features
* Compares two Excel files and highlights the differences in red
* Supports .xlsx and .xls file formats
* Easy to use with Telegram's user-friendly interface
### Requirements
The following packages are required to run the bot:

* python-telegram-bot
* openpyxl
* xlrd
## Installation
Clone the repository and install the required packages:

```bash
git clone https://github.com/vanchaxy/excel-comparison-bot.git
cd excel-comparison-bot
pip install -r requirements.txt
```
Create a .env file in the root of the project and set your Telegram bot token as the value of TELEGRAM_BOT_TOKEN:

```makefile
TELEGRAM_BOT_TOKEN=your_token_here
```
Run the bot:
```
python bot.py
```
## Contribute
This bot was created for educational purposes and there is always room for improvement. If you have any suggestions or improvements, feel free to fork the repository and submit a pull request.
