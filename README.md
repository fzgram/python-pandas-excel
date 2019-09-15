# python-pandas-excel-automation

This project does a small scale data analysis of a file which consists of consolidation results of the 200 set of FPS (Frames Per Second) values reported by an Augmented Reality (AR) glasses. It analyzes the input file, filters out outliers and non outliers values (in this example, results below 120 FPS are outliers). Additionally, it does basic data analysis of the FPS values by providing min, max, average, standard deviation and distribution for each FPS value reported, with a piechart to depict the information.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

- Install `python 3.6.8` or higher. 
   1. Windows
      - [Click here](https://www.howtogeek.com/197947/how-to-install-python-on-windows/) and follow section **"How to Install Python 3"**
      - ***`OPTIONAL`*** Install `git bash` for UNIX like command line experience on Windows. [Click here](https://www.techoism.com/how-to-install-git-bash-on-windows/) for instructions
      
   2. Unix (Ubuntu, Cent OS, etc.)
      - [Click here](https://www.tecmint.com/install-python-in-linux/) and follow instructions

   **Make sure `python` is added to your PATH for direct access from the terminal**. Confirm installation by going to the terminal and      type `python`. it should start an interactive python. Example below
   
   ```
   C:\>python
   Python 3.6.0 (v3.6.0:41df79263a11, Dec 23 2016, 07:18:10) [MSC v.1900 32 bit (Intel)] on win32
   Type "help", "copyright", "credits" or "license" for more information.
   >>>
   ```
 
 - Install python `virtualenv` package
   1. Go to terminal of your OS (or git bash on Windows)
   2. Use command to install virtualenv `pip install virtualenv`
   
 - Clone this project on your local machine
   1. Create an empty directory in your local machine (Eg: C:\new_project)
   2. `cd` into the directory.
   3. Click on "Clone or download" tab on top right corner of this page, copy the URL
   4. Go back to terminal and run the command `git clone <COPIED URL HERE>`
   
 - Create virtual environment for the project
   1. `cd` into top location of the project (for eg: C:\new_project\python-pandas-excel-automation)
   2. Run the command `virtualenv py3`. This will create a virtual environment directory `py3`.
   3. Activate the virtual environment
      - UNIX    => `source py3/Scripts/activate`
      - WINDOWS => `py3\Scripts\activate.bat`
      After activation, you should see a `(py3)` on the command prompt.
      
      Example (from Command Prompt)
      ```
      D:\my_projects\python-pandas-excel-automation>py3\Scripts\activate.bat
      
      (venvpy3) D:\projects\python-pandas-excel-automation>
      ```
      
      Example from `git bash`
      ```
      Shardulito@DESKTOP MINGW64 /d/projects/python-pandas-excel-automation (master)
      $ source py3/Scripts/activate
      
      (venvpy3)
      Shardulito@DESKTOP MINGW64 /d/projects/python-pandas-excel-automation (master)
      $
      ```

## Running the tests

You can either run the script using jupyter lab or directly running the python version

- Run using jupyter
  1. Type `jupyter lab` on your terminal, it will open a web browser tab with the code.
  2. Click on `Kernel` -> `Restart Kernel and Run All Cells`
  3. Look for output in folder `output`
  
- Run using python
  - WINDOWS
    - Use command `python ar_fps_post_statistics.py input\consolidation_result_ARGlass_TypeA.xlsx`
  - UNIX (or git bash)
    - Use command `python ar_fps_post_statistics.py input/consolidation_result_ARGlass_TypeA.xlsx`


