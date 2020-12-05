h0 = figure;
set(h0,'unit','normalized','color','default',...
    'position',[0.1 0.1 0.8 0.8],'menu','none',...
    'numbertitl','off','name','GUI 生成菜单');
h = uicontextmenu;       %创建右键菜单
set(h0,'uicontextmenu',h);%将右键菜单和窗口相联系设置窗口uicontextmenu属性值是h
h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
% h1 = uimenu(h,'label','一级菜单');
%%
h2 = uimenu(h1,'label','二级菜单');
h3 = uimenu(h2,'label','三级菜单');
h4 = uimenu(h3,'label','四级菜单');
h5 = uimenu(h4,'label','五级菜单');
h6 = uimenu(h5,'label','六级菜单');
h7 = uimenu(h6,'label','七级菜单');
h8 = uimenu(h7,'label','八级菜单');
h9 = uimenu(h8,'label','九级菜单');