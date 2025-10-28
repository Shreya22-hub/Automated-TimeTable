# run_both.py
from multiprocessing import Process
import subprocess

def run_app1():
    subprocess.run(["python", "app.py"])

def run_app2():
    subprocess.run(["python", "app2.py"])

if __name__ == "__main__":
    p1 = Process(target=run_app1)
    p2 = Process(target=run_app2)
    p1.start()
    p2.start()
    p1.join()
    p2.join()