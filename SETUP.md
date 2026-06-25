# Setup Guide

How to get **EasyPower Reporting Automation** running on a Windows PC — no programming experience needed. Do these once.

## 1. Install Python
1. Go to **https://www.python.org/downloads/** and click the big yellow **Download Python** button.
2. Run the downloaded installer. On the very first screen, **check the box "Add python.exe to PATH"** (it's at the bottom and easy to miss — but essential), then click **Install Now**.
   - ⚠️ Do **not** install "Python" from the Microsoft Store — that version will not work here.
3. Check it worked: open the Start menu, type **cmd**, open **Command Prompt**, type `python --version`, press Enter. You should see something like `Python 3.12.4`. (If it says "not found", re-run the installer and make sure the PATH box is checked.)

## 2. Get the project files
If you already have the project folder, skip this. Otherwise, on the project's GitHub page click the green **Code** button → **Download ZIP**, then unzip it somewhere easy like your **Documents** folder.

## 3. Open a terminal inside the project folder
1. Open the project folder in File Explorer.
2. Click the **address bar** (the strip showing the folder path), type **cmd**, and press Enter. A black Command Prompt window opens, already pointed at the folder.

## 4. Install the required packages
In that Command Prompt, type this and press Enter (you need internet):
```
python -m pip install -r requirements.txt
```
Wait a minute or two for it to finish.

## 5. Open EasyPower
- Launch EasyPower and open your study's **Base Case** `.dez` file.
- Set your **print settings** and **report settings** the way you want them — the program will pause and ask you to confirm these.
- *Tip:* work on a **copy** of your study file if you don't want scenario results saved back into the original.

## 6. Run it
In the Command Prompt, type:
```
python easypower_batch_reports.py
```
Then follow the on-screen prompts:
- **Output directory** — where to save the reports (or press Enter for the default).
- **Scenarios** — press Enter to run all of them, or type specific names separated by commas (e.g. `Scenario-1, Scenario-5`).
- It asks you to confirm EasyPower is open and your settings are set — type **yes** to each.

It then runs on its own. **Don't touch the mouse or keyboard while it runs** — it's driving the screen for you. When it finishes, your output folder will hold the individual reports plus **Combined_Reports.xlsx** and **Combined_Reports.pdf**.

---

Stuck? Email **cooper@powerlinktec.com**.
