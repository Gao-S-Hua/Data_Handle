import os
import shutil

def prepare():
  if os.path.exists(r'./outputs'):
    try:
      shutil.rmtree(r'./outputs')
      print("| Removed old directory and files")
    except:
      print("** ERROR: Some file is locked, please close it!")
      exit(0)
  os.mkdir(r'./outputs')
  os.mkdir(r'./outputs/img')
  print("| Created new directory")

if __name__ == "__main__":
  prepare()
