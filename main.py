import math
import msvcrt
import sys
import time
from pathlib import Path
from openpyxl import load_workbook

FILE_PATH = "Spreadsheet Jun.xlsx"


def to_float(value, default=None):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def clock(seconds):
    h, rem = divmod(int(seconds), 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def load_tasks(path):
    try:
        wb = load_workbook(path)
    except FileNotFoundError:
        print(f"Could not find '{path}'.")
        return []

    tasks = []
    rows = wb.active.iter_rows(min_row=2, max_col=3, values_only=True)
    for row_num, (name, budget, tracked) in enumerate(rows, start=2):
        budget = to_float(budget)
        if not name or not str(name).strip() or not budget or budget <= 0:
            continue
        tasks.append({
            "row": row_num,
            "task": str(name).strip(),
            "budgeted": budget,
            "tracked": to_float(tracked, 0.0),
        })

    wb.close()
    return tasks


def save_tasks(path, tasks):
    try:
        wb = load_workbook(path)
        sheet = wb.active
        sheet.cell(row=1, column=3).value = "Hours Tracked"
        for t in tasks:
            sheet.cell(row=t["row"], column=3).value = t["tracked"]
        wb.save(path)
        wb.close()
        print("Saved.")
    except PermissionError:
        print("Couldn't save. Is the Excel file still open?")
    except Exception as e:
        print(f"Couldn't save: {e}")


def run_timer(task):
    print(f"\nWorking on: {task['task']}")
    input("Press ENTER to start...")

    start = time.time()
    print("Timer running. Press ENTER to stop.")

    while True:
        sys.stdout.write(f"\rElapsed: {clock(time.time() - start)}")
        sys.stdout.flush()
        if msvcrt.kbhit() and msvcrt.getch() == b"\r":
            break
        time.sleep(0.1)

    hours = (time.time() - start) / 3600
    print(f"\nTime worked: {hours:.2f} hours")
    return hours


def adjust_hours(task):
    print(f"\nCurrently tracked: {task['tracked']:.2f} hours")

    while True:
        raw = input("Hours to add or subtract (0 to cancel): ")
        amount = to_float(raw)

        if amount is None or not math.isfinite(amount):
            print("Enter a valid number.")
            continue
        if amount == 0:
            return False
        if task["tracked"] + amount < 0:
            print("Tracked hours can't go below zero.")
            continue

        task["tracked"] += amount
        return True


def task_menu(task):
    """Returns True if the task's hours changed."""
    while True:
        print(f"\n=== {task['task']} ===")
        print(f"Tracked:  {task['tracked']:.2f}")
        print(f"Budgeted: {task['budgeted']:.2f}")
        print("\n1. Start timer\n2. Adjust hours\n3. Back")

        choice = input("> ").strip()
        if choice == "1":
            task["tracked"] += run_timer(task)
            return True
        if choice == "2":
            return adjust_hours(task)
        if choice == "3":
            return False
        print("Pick 1, 2, or 3.")


def select_task(tasks):
    print("\n=== Tasks ===")
    for i, t in enumerate(tasks, start=1):
        print(f"{i}. {t['task']} ({t['tracked']:.2f} / {t['budgeted']:.2f} hrs)")
    print("0. Exit")

    while True:
        try:
            choice = int(input("> "))
        except ValueError:
            print("Enter a number.")
            continue

        if choice == 0:
            return None
        if 1 <= choice <= len(tasks):
            return tasks[choice - 1]
        print(f"Enter a number from 0 to {len(tasks)}.")


def main():
    tasks = load_tasks(FILE_PATH)
    if not tasks:
        print("No valid tasks found.")
        return

    while True:
        task = select_task(tasks)
        if task is None:
            break

        if task_menu(task):
            print(f"{task['task']}: {task['tracked']:.2f} hours tracked.")
            save_tasks(FILE_PATH, tasks)


if __name__ == "__main__":
    main()