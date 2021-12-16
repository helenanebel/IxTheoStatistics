from cx_Freeze import setup, Executable

setup(name = "IxTheo_Statistik" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("ixtheo_statistics.py")])

# C:\Users\hnebel\PycharmProjects\IxTheoStatistics>python setup.py build