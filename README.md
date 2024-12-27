# Project Title

## Project Goal

The goal of this project is to automate the process of organizing job folders on a network fileshare based on input from a schedule spreadsheet. The script `job_organizer.py` creates or modifies existing files within specified job folders. Future features will include a script that creates a material required list based on data provided from .kss files and imports it into a spreadsheet that can be manipulated after the fact.

## Using the pathlib Library

The `pathlib` library in Python is used extensively in this project for handling filesystem paths. Here's a quick reminder on how to use it:

### Creating a Path

You can create a new path using the `Path` class:

### pathlib Overview
from pathlib import Path

p = Path('/path/to/directory')

Joining Paths
You can join paths using the / operator:

p = Path('/path/to/directory')
file_path = p / 'file.txt'

Checking if a Path Exists
You can check if a path exists using the exists() method:

if p.exists():
    print("Path exists")

Creating a Directory
You can create a new directory using the mkdir() method:

p.mkdir(parents=True, exist_ok=True)

The parents=True argument means that any missing parent directories will be created as well. The exist_ok=True argument means that no error will be raised if the directory already exists.

##  Future Features
