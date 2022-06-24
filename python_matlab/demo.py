# coding:utf-8
# 1 在matlab命令行输入matlabroot 找到maltab路径 matlab 2018b python35
# 2 以管理员运行cmd 进入maltab路径\extern\engines\python
# 3 python setup.py install
# 参考https://zhuanlan.zhihu.com/p/392728491
import matlab
import matlab.engine               # 在Python中导入Matlab引擎
path = 'D:/File_Pycharm/py35_prj/python_matlab'
eng = matlab.engine.start_matlab() # 启动Matlab引擎，命名为eng
eng.cd(path,nargout=0)
# eng.ini_operating_point(nargout=0) # 通过eng运行写好的m文件 “ini_operating_point.m”，nargout=0表示不返回输出。
#                                    # Note: 此m文件需要在ptyhon的启动界面下。、
# Q_ref='sddd'
# out = eng.eval('Inport')  # 用 matlab的eval 函数读取变量 ‘out’
# eng.workspace['Iport'] = list(Q_ref)  # 在Matlab工作区生成变量‘Q_ref’，大小等于Python中的变量‘Q_ref’
o=[]
# eng.getVariable('Inport')
# eng.workspace['Inport']
print(o)
# out = eng.eval('out')  # 用 matlab的eval 函数读取变量 ‘out’
# print(out)
inport,outport,constant  = eng.FindPortdemo(nargout=3)
# out = eng.eval('Inport')
print("inport",inport)
print("outport",outport)
print("constant",constant)
eng.exit()
