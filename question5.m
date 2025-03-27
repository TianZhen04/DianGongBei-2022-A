% 定义机组1的数据
TP1 = struct('MaxOutput', 600, 'MinOutput', 180, 'CO2Emission', 720, ...
    'C', 786.80, 'B', 30.42, 'A', 0.226);
TP2 = struct('MaxOutput', 300, 'MinOutput', 90, 'CO2Emission', 750, ...
    'C', 451.32, 'B', 65.12, 'A', 0.588);
TP3 = struct('MaxOutput', 150, 'MinOutput', 45, 'CO2Emission', 790, ...
    'C', 1049.50, 'B', 139.6, 'A', 0.785);

WP.price = 45;
TP.price = 0.7;

Bat = struct('PowerCost', 3000, 'EnergyCost', 3000, ...
    'OperationalCost', 0.05);

data = readmatrix("附件1.xlsx",'Range', 'B2:H97');

Pload = data(:,3);
Pwind300 = data(:,4);
Pwind600 = data(:,5);
Pwind900 = data(:,6);

%%
% 第五问
P1 = sdpvar(24*4, 1, 'full');
W900 = sdpvar(24*4, 1, 'full');
W900loss = sdpvar(24*4, 1, 'full');
Ploadloss = sdpvar(24*4, 1, 'full');

cons = [P1 <= TP1.MaxOutput];
cons = [cons, P1 >= TP1.MinOutput];
cons = [cons, W900 >= 0, W900loss>=0, Ploadloss>=0];
cons = [cons, W900+W900loss == Pwind900];
cons = [cons, P1+W900+Ploadloss == Pload];

A = 80;
cost = cal_TPcost(P1, TP1, A)+cal_Wcost(W900, W900loss)+sum(Ploadloss)*8000;

% 进行优化
ops = sdpsettings('solver', 'gurobi');
result = optimize(cons, cost, ops);

P1 = value(P1);
W900 = value(W900);
W900loss = value(W900loss);
Ploadloss = value(Ploadloss);
% 绘制堆叠柱状图
figure
% 注意：堆叠柱状图的输入需要是矩阵形式
bar_data = [P1, W900, -W900loss,Ploadloss]'; % 将数据转置为矩阵
hold on
bar(0:0.25:23.75, bar_data, 'stacked');
plot(0:0.25:23.75, Pload, 'LineWidth', 2, 'Color', [0.5, 0.5, 0.5]); % 灰色曲线表示负荷
% 美化图形
xlabel('小时', 'FontSize', 12);
ylabel('功率 (单位：kW)', 'FontSize', 12);
title('24小时负荷与发电分布', 'FontSize', 14);

% 设置图例
legend( '机组1', '900MW风电', '风电弃电','失负荷','负荷', ...
    'Location', 'northeast', 'FontSize', 10);

%%
% 第五问
P1 = sdpvar(24*4, 1, 'full');
W900 = sdpvar(24*4, 1, 'full');
W900loss = sdpvar(24*4, 1, 'full');

Pinbat = sdpvar(24*4,1,'full');
Poutbat = sdpvar(24*4,1,'full');
Wbat = sdpvar(24*4+1,1,'full');

BatPN = sdpvar(1, 1, 'full');
BatWN = sdpvar(1, 1, 'full');

% 上下限
cons = [P1 <= TP1.MaxOutput];
cons = [cons, P1 >= TP1.MinOutput];
cons = [cons, W900 >= 0, W900loss>=0];
% 风电平衡
cons = [cons, W900+W900loss == Pwind900];
% 供需平衡
cons = [cons, P1+W900+Poutbat*0.9 == Pload + Pinbat/0.9];
% 储能约束
cons = [cons, diff(Wbat)==(Pinbat-Poutbat)*0.25];
cons = [cons, Wbat<=BatWN, Wbat>=0,Wbat(1) == 0];
cons = [cons, Pinbat>=0, Pinbat<=BatPN, Poutbat>=0, Poutbat<=BatPN];
A = 80;
% cost = cal_TPcost(P1, TP1, A)+cal_Wcost(W900,W900loss)+sum(W900loss)*300+cal_Batcost(BatPN,BatWN,Pinbat,Poutbat);
cost = BatWN;%储能容量最小

% 进行优化
ops = sdpsettings('solver', 'gurobi');
result = optimize(cons, cost, ops);

P1 = value(P1);
W900 = value(W900);
W900loss = value(W900loss);
Pinbat = value(Pinbat);
Poutbat = value(Poutbat);
Wbat = value(Wbat);
BatPN = value(BatPN);
BatWN = value(BatWN);

% 绘制堆叠柱状图
figure
% 注意：堆叠柱状图的输入需要是矩阵形式
bar_data = [P1, W900, -W900loss,Poutbat*0.9]'; % 将数据转置为矩阵
hold on
bar(0:0.25:23.75, bar_data, 'stacked');
plot(0:0.25:23.75, Pload, 'LineWidth', 2, 'Color', [0.5, 0.5, 0.5]); % 灰色曲线表示负荷
plot(0:0.25:23.75, Pload+ Pinbat/0.9, 'LineWidth', 2); 
% 美化图形
xlabel('小时', 'FontSize', 12);
ylabel('功率 (单位：kW)', 'FontSize', 12);
title('24小时负荷与发电分布', 'FontSize', 14);

% 设置图例
legend( '机组1', '900MW风电', '风电弃电','储能放电','负荷','负荷和储能充电', ...
    'Location', 'northeast', 'FontSize', 10);

figure;
yyaxis left;
bar_data = [-Poutbat+Pinbat]'; % 将数据转置为矩阵
bar(0:0.25:23.75, bar_data, 'stacked');
ylabel('充电功率 (W)');
ylim([-200 500]); % 设置y轴范围
yyaxis right;
plot(0:0.25:24, Wbat, 'b', 'LineWidth', 2); % 假设Wbat是储能容量数据
ylabel('储能容量 (%)');
ylim([-200 500]); % 设置y轴范围
% 设置x轴标签和图形标题
xlabel('时间 (小时)');
title('电池充放电功率与储能容量');
% 添加图例
legend('充放电功率', '储能容量', 'Location', 'best');

% 显示网格
grid on;

% 优化图形外观
box on; % 添加边框

% 设置字体以支持中文显示
set(gca, 'FontName', 'SimHei'); % 设置当前坐标轴字体为黑体
set(findobj(gca,'type','text'),'FontName','SimHei'); % 设置坐标轴内所有文本的字体为黑体

function cost = cal_TPcost(P, TP, A)
    price = 0.7;
    A = A/1000;
    F = TP.A * P.^2 + TP.B * P + TP.C;% 煤耗量
    F = F*0.25;
    F = sum(F);
    cost1 = price * F;% 煤价
    cost2 = cost1 * 0.5;% 运行维护成本
    cost3 = A * F;% 碳捕集成本
    cost = cost1+cost2+cost3;
end

function cost = cal_Wcost(W,Wloss)
    price_W = 45;price_Wloss = 300;
    W_sum = sum(W) * 0.25;
    Wloss_sum = sum(Wloss) * 0.25;
    cost = W_sum * price_W + Wloss_sum * (price_W + price_Wloss);
end

function cost = cal_Batcost(PN,WN,Pin,Pout)
    cost1 = 3000000*PN + 3000000*WN;
    cost2 = sum(Pin+Pout)*0.25*50;
    cost = cost1/10/365 + cost2;
end
