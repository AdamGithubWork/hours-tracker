from openpyxl import load_workbook
import time
import sys
import msvcrt


FILE_PATH = "Spreadsheet Jun.xlsx"


def load_tasks(file_path):
    tasks = []

    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active


        if sheet.cell(row=1, column=3).value is None:
            sheet.cell(row=1, column=3).value = "Hours Tracked"

        for row in range(2, sheet.max_row + 1):
            task = sheet.cell(row=row, column=1).value
            hours_budgeted = sheet.cell(row=row, column=2).value
            hours_tracked = sheet.cell(row=row, column=3).value


            if task is None or str(task).strip() == "":
                continue


            if hours_budgeted is None or str(hours_budgeted).strip() == "":
                continue

            try:
                hours_budgeted = float(hours_budgeted)
            except (ValueError, TypeError):
                continue


            if hours_budgeted <= 0:
                continue


            try:
                hours_tracked = float(hours_tracked)
            except (ValueError, TypeError):
                hours_tracked = 0.0

            tasks.append({
                "row": row,
                "task": str(task).strip(),
                "hours_budgeted": hours_budgeted,
                "hours_tracked": hours_tracked
            })

        workbook.close()

    except FileNotFoundError:
        print(f"Could not find '{file_path}'.")
        return []

    return tasks


def save_tasks(file_path, tasks):
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active

        sheet.cell(row=1, column=3).value = "Hours Tracked"

        for task in tasks:
            sheet.cell(
                row=task["row"],
                column=3
            ).value = task["hours_tracked"]

        workbook.save(file_path)
        workbook.close()

        print("\nHours saved successfully.")

    except PermissionError:
        print(
            "\nCould not save the file. "
            "Make sure the Excel file is closed."
        )

    except Exception as e:
        print(f"\nCould not save the file: {e}")


def select_task(tasks):
    print("\n=== Tasks ===")

    for i, task in enumerate(tasks, start=1):
        print(
            f"{i}. {task['task']} "
            f"({task['hours_tracked']:.2f} / "
            f"{task['hours_budgeted']:.2f} hours)"
        )

    print("0. Exit")

    while True:
        choice = input("\nSelect a task: ")

        try:
            choice = int(choice)

            if choice == 0:
                return None

            if 1 <= choice <= len(tasks):
                selected_task = tasks[choice - 1]

                print(f"\nSelected: {selected_task['task']}")

                return selected_task

            print(
                f"Please enter a number between "
                f"0 and {len(tasks)}."
            )

        except ValueError:
            print("Please enter a number.")


def run_timer(task):
    print(f"\nWorking on: {task['task']}")
    input("Press ENTER to start the timer...")

    start_time = time.time()

    print("\nTimer started.")
    print("Press ENTER to stop the timer.")

    while True:
        elapsed_seconds = time.time() - start_time

        hours = int(elapsed_seconds // 3600)
        minutes = int((elapsed_seconds % 3600) // 60)
        seconds = int(elapsed_seconds % 60)

        sys.stdout.write(
            f"\rElapsed: {hours:02d}:{minutes:02d}:{seconds:02d}"
        )
        sys.stdout.flush()

        if msvcrt.kbhit():
            key = msvcrt.getch()

            if key == b'\r':
                break

        time.sleep(0.1)

    end_time = time.time()

    elapsed_seconds = end_time - start_time
    elapsed_hours = elapsed_seconds / 3600

    hours = int(elapsed_seconds // 3600)
    minutes = int((elapsed_seconds % 3600) // 60)
    seconds = int(elapsed_seconds % 60)

    print()
    print(f"Time worked: {elapsed_hours:.2f} hours")

    return elapsed_hours


def main():
    tasks = load_tasks(FILE_PATH)

    if not tasks:
        print("No valid tasks found.")
        return

    while True:
        selected_task = select_task(tasks)

        if selected_task is None:
            print("\nGoodbye!")
            break

        hours_worked = run_timer(selected_task)

        selected_task["hours_tracked"] += hours_worked

        print(
            f"{selected_task['task']} now has "
            f"{selected_task['hours_tracked']:.2f} "
            f"hours tracked."
        )

        save_tasks(FILE_PATH, tasks)


if __name__ == "__main__":
    main()