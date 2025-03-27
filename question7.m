data = readmatrix('附件2.xlsx','Range','C2:D1441');
Pload = data(:,1);
PW = data(:,2);
% 定义机组1的数据
TP1 = struct('MaxOutput', 600, 'MinOutput', 180, 'CO2Emission', 720, ...
    'C', 786.80, 'B', 30.42, 'A', 0.226);
TP2 = struct('MaxOutput', 300, 'MinOutput', 90, 'CO2Emission', 750, ...
    'C', 451.32, 'B', 65.12, 'A', 0.588);
TP3 = struct('MaxOutput', 150, 'MinOutput', 45, 'CO2Emission', 790, ...
    'C', 1049.50, 'B', 139.6, 'A', 0.785);

Bat = struct('PowerCost', 3000, 'EnergyCost', 3000, ...
    'OperationalCost', 0.05);

%%
%不使用储能调节
P1 = sdpvar(24*4*15, 1, 'full');
W = sdpvar(24*4*15, 1, 'full');
Wloss = sdpvar(24*4*15, 1, 'full');
Ploadloss = sdpvar(24*4*15, 1, 'full');

cons = [P1 <= TP1.MaxOutput];
cons = [cons, P1 >= TP1.MinOutput];
cons = [cons, W >= 0, Wloss>=0, Ploadloss>=0];
cons = [cons, W+Wloss == PW];
cons = [cons, P1+W+Ploadloss == Pload];

A = 80;
cost = cal_TPcost(P1, TP1, A)+cal_Wcost(W, Wloss)+sum(Ploadloss)*8000;

% 进行优化
ops = sdpsettings('solver', 'gurobi');
result = optimize(cons, cost, ops);

P1 = value(P1);
W = value(W);
Wloss = value(Wloss);
Ploadloss = value(Ploadloss);
% 绘制堆叠柱状图
figure
% 注意：堆叠柱状图的输入需要是矩阵形式
bar_data = [P1, W, -Wloss,Ploadloss]'; % 将数据转置为矩阵
hold on
bar(0:0.25:24*15-0.25, bar_data, 'stacked');
plot(0:0.25:24*15-0.25, Pload, 'LineWidth', 2, 'Color', [0.5, 0.5, 0.5]); % 灰色曲线表示负荷
% 美化图形
xlabel('小时', 'FontSize', 12);
ylabel('功率 (单位：kW)', 'FontSize', 12);
title('24小时负荷与发电分布', 'FontSize', 14);

% 设置图例
legend( '机组1出力', '风电出力', '风电弃电','失负荷','负荷', ...
    'Location', 'northeast', 'FontSize', 10);

%%
%使用储能调节
P1 = sdpvar(24*4*15, 1, 'full');
W = sdpvar(24*4*15, 1, 'full');
Wloss = sdpvar(24*4*15, 1, 'full');
Ploadloss = sdpvar(24*4*15, 1, 'full');

Pinbat = sdpvar(24*4*15,1,'full');
Poutbat = sdpvar(24*4*15,1,'full');
Wbat = sdpvar(24*4*15+1,1,'full');

BatPN = sdpvar(1, 1, 'full');
BatWN = sdpvar(1, 1, 'full');


cons = [P1 <= TP1.MaxOutput];
cons = [cons, P1 >= TP1.MinOutput];
cons = [cons, W >= 0, Wloss>=0, Ploadloss>=0];
cons = [cons, W+Wloss == PW];

cons = [cons, P1+W+Poutbat*0.9+Ploadloss == Pload + Pinbat/0.9];
cons = [cons, diff(Wbat)==(Pinbat-Poutbat)*0.25];
cons = [cons, Wbat<=BatWN, Wbat>=0,Wbat(1) == 0];
cons = [cons, Pinbat>=0, Pinbat<=BatPN, Poutbat>=0, Poutbat<=BatPN];

A = 80;
% cost = cal_TPcost(P1, TP1, A)+cal_Wcost(W, Wloss)+sum(Ploadloss)*8000;
cost = cal_TPcost(P1, TP1, A)+cal_Wcost(W,Wloss)+sum(Wloss)*300+cal_Batcost(BatPN,BatWN,Pinbat,Poutbat)+sum(Ploadloss)*8000;

% 进行优化
ops = sdpsettings('solver', 'gurobi');
result = optimize(cons, cost, ops);

P1 = value(P1);
W = value(W);
Wloss = value(Wloss);
Ploadloss = value(Ploadloss);

Pinbat = value(Pinbat);
Poutbat = value(Poutbat);
Wbat = value(Wbat);
BatPN = value(BatPN);
BatWN = value(BatWN);



% 绘制堆叠柱状图
figure
% 注意：堆叠柱状图的输入需要是矩阵形式
bar_data = [P1, W, -Wloss,Poutbat*0.9,Ploadloss]'; % 将数据转置为矩阵
hold on
bar(0:0.25:24*15-0.25, bar_data, 'stacked');
plot(0:0.25:24*15-0.25, Pload, 'LineWidth', 2, 'Color', [0.5, 0.5, 0.5]); % 灰色曲线表示负荷
plot(0:0.25:24*15-0.25, Pload+ Pinbat/0.9, 'LineWidth', 2); 
% 美化图形
xlabel('小时', 'FontSize', 12);
ylabel('功率 (单位：kW)', 'FontSize', 12);
title('24小时负荷与发电分布', 'FontSize', 14);

% 设置图例
legend( '机组1出力', '风电出力', '风电弃电','储能放电','失负荷','负荷','负荷和储能充电', ...
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
