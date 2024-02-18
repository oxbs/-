import os

print(os.path.abspath(__file__))
print(os.path.dirname(os.path.abspath(__file__)))
dirdir = os.path.dirname(os.path.abspath(__file__))
os.path.join(dirdir,'html','index.html')
print(os.path.join(dirdir,'html','index.html'))

