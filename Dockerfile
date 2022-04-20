# syntax=docker/dockerfile:1
FROM python:3.9.7
ADD newMachine.py /
RUN pip install pyautogui
RUN pip install pandas
CMD ["python", "newMachine.py"]