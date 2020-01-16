import os

root = r"C:\Users\sb512911\Downloads\VDR data"
list1 = []
for path, subdirs, files in os.walk(root):
    for name in files:
        a = os.path.join(path, name).split('.')
        list1.append(a[-1])
print(set(list1))

