from distutils.core import setup
import py2exe

setup(console=['evade.py'],
      options={
          "py2exe": {
              "packages": ["win32com", "webbrowser", "requests", "time", "os", "shutil", "bs4", "smtplib",
                           "email", "re", "tkinter", "random", "threading"],
          }
      },
      zipfile=None)