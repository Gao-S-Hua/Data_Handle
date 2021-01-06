import os

class TestResult:
  def __init__(self, SN, temperature):
    self.SN = SN
    self.temperature = temperature
    self.data = []

def get_path(path):
  return os.getcwd() + path