from tools import *
from models import *

if __name__ == '__main__':
    gather.gath_data('src',Stud,('seq','name','addr'))
    xlstools.dump(Stud,('序号','姓名','地址'),('seq','name','addr'),'addr')