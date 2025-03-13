# Currency Exchange rate data Bot for PoE2

## Description

This application is a bot that automates actions for viewing currency exchange rates in the game Path of Exile 2 (PoE2). It takes control of the mouse and captures screenshots of each currency pair. The currency pairs are generated from the `items.txt` file. In the future, it is planned to analyze screenshots and extract text with exchange rates from them.

**Warning:** By using this application, the user assumes all risks. The application violates the agreement with the developer of the game Path of Exile 2 (PoE2). The account may be banned!

## Requirements

- The game must be in English (localization is planned for the future)
- Python 3.x
- Libraries: pyautogui, pillow, keyboard, pywin32

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/tautau-code/rateScramblerBot
   ```
2. Navigate to the project directory:
   ```
   cd rateScramblerBot
   ```
3. Install the required dependencies:
   ```
   pip install pyautogui Pillow keyboard pywin32
   ```

## Usage

1. Ensure the `items.txt` file is located in the root directory of the project and contains a list of items to process.
2. Open the game and the Currency Exchange window.
3. Run the script:
   ```
   python main.py
   ```
4. Confirm the start of the script by typing "yes" in the console.
5. Press the Backspace key to stop the script.
6. The application processes item pairs quite slowly to allow the game to load everything, so information gathering may take a long time. To avoid starting over each time the application is run, it has a continuation logic (which will be corrected as it currently works incorrectly by checking the name of the last pair from the screenshots).

## Author

Lev Efimenko