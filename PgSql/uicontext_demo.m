h0 = figure;
set(h0,'unit','normalized','color','default',...
    'position',[0.1 0.1 0.8 0.8],'menu','none',...
    'numbertitl','off','name','GUI ���ɲ˵�');
h = uicontextmenu;       %�����Ҽ��˵�
set(h0,'uicontextmenu',h);%���Ҽ��˵��ʹ�������ϵ���ô���uicontextmenu����ֵ��h
h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
% h1 = uimenu(h,'label','һ���˵�');
%%
h2 = uimenu(h1,'label','�����˵�');
h3 = uimenu(h2,'label','�����˵�');
h4 = uimenu(h3,'label','�ļ��˵�');
h5 = uimenu(h4,'label','�弶�˵�');
h6 = uimenu(h5,'label','�����˵�');
h7 = uimenu(h6,'label','�߼��˵�');
h8 = uimenu(h7,'label','�˼��˵�');
h9 = uimenu(h8,'label','�ż��˵�');