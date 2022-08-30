# Workbook model

This code is used to copy data between 2 excel documents. Data is copied from source file and pasted to corresponding sheet in template file.

# App model

This code is used to model a graphical user interface. It lets to set filepaths to template file, source file and folder where newly created file will be saved. It lets copy data between 2 excel files.

## Requirements

- Conda: click [here](https://docs.conda.io/projects/conda/en/latest/user-guide/install/index.html) for the installation guide 
- Libraries: the required libraries are listed in the _'environment.yml'_ file and can be installed from the command line in the project directory with the following command: ```conda env create -f environment.yml```  

## Project organization

```
├── src
│   └── gui.py              <- Contains an App class with all model logic
│   └── utilis.py           <- Contains helper functions
│   └── workbook.py         <- Contains a Workbook class with all model logic
├── environment.yml             <- Required Python libraries to run the code
├── main.py                     <- CLI script to execute the App model
└── README.md                   <- This project description
```

## Run code

The model can be executed from the command line by running the following command:

```
python main.py