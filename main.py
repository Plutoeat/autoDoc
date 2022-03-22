import sys

from model.saveModel import savedoc

if __name__ == '__main__':
    title = sys.argv[1]
    doctype = sys.argv[2]
    if doctype == "论文":
        savedoc(title, "paper")
