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
% 第一问
P1 = sdpvar(24*4, 1, 'full');
P2 = sdpvar(24*4, 1, 'full');
P3 = sdpvar(24*4, 1, 'full');

% 出力上下限
cons = [P1 <= TP1.MaxOutput, P2 <= TP2.MaxOutput, P3 <= TP3.MaxOutput];
cons = [cons, P1 >= TP1.MinOutput, P2 >= TP2.MinOutput, P3 >= TP3.MinOutput];
% 供需平衡
cons = [cons, P1+P2+P3 == Pload];

A = 80;
cost = cal_TPcost(P1, TP1, A)+cal_TPcost(P2, TP2, A)+cal_TPcost(P3, TP3, A);

% 进行优化
ops = sdpsettings('solver', 'gurobi');
result = optimize(cons, cost, ops);

P1 = value(P1);
P2 = value(P2);
P3 = value(P3);

% 绘制堆叠柱状图
% 注意：堆叠柱状图的输入需要是矩阵形式
bar_data = [P1, P2, P3]'; % 将数据转置为矩阵
hold on
bar(0:0.25:23.75, bar_data, 'stacked');
plot(0:0.25:23.75, Pload, 'LineWidth', 2, 'Color', [0.5, 0.5, 0.5]); % 灰色曲线表示负荷
% 美化图形
xlabel('小时', 'FontSize', 12);
ylabel('功率 (单位：kW)', 'FontSize', 12);
title('24小时负荷与发电分布', 'FontSize', 14);

% 设置图例
legend( '机组1', '机组2', '机组3','负荷', ...
    'Location', 'northeast', 'FontSize', 10);


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
