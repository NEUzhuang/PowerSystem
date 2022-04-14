clc
clear
Bus=xlsread('bus.xlsx');
Branch=xlsread('branch.xlsx');
[busnum,~]=size(Bus);
[branchnum,row]=size(Branch);
soubus=Branch(:,2);
mobus=Branch(:,3);
Vbus=ones(busnum,1);
Vbus1=Vbus;
Ploss=zeros(busnum,1);
Qloss=zeros(busnum,1);
e=1;
i=1;
k=0;
Branch1=Branch;
n=1;
%% 精髓
%%%%%%%%%%%%%%%支路重新排序，各个分支线同时进行计算
while ~isempty(Branch1)%%%%T1：支路矩阵。
    m=1;
    [s,row]=size(Branch1);
        while s>0
            t=find(Branch1(:,2)== Branch1(s,3));%判断是否是子节点
                if isempty(t) %子节点
                    T1(n,:)= Branch1(s,:);%从节点系统末端向首端进行；
                    n=n+1;
                else
                    T2(m,:)= Branch1(s,:);%不是子节点
                     m=m+1;
                end
               s=s-1;
        end
        Branch1=T2;
        T2=[];
end
%% 
%%%%%%%%%%%%%%%%%%%%%%%
while e>1.0e-07%精度
    %%%%%%%%%%%%%%%%%%%%%%%%从末端向首端推功率%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    P=zeros(busnum,1);%%存放功率
    Q=zeros(busnum,1);
        for s=1:branchnum
                i=T1(s,2);
                j=T1(s,3);
                R=T1(s,4);
                X=T1(s,5);
                Pload=Bus(j,2);%有功负荷
                Qload=Bus(j,3);%无功负荷
                II=((Pload+P(j))^2+(Qload+Q(j))^2)/(Vbus(j)^2*1000);
                Ploss(i,j)=II*R;%支路有功损耗
                Qloss(i,j)=II*X;%支路无功损耗
                P(i,j)=Pload+Ploss(i,j)+P(j);
                Q(i,j)=Qload+Qloss(i,j)+Q(j);
                P(i)=P(i)+P(i,j);
                Q(i)=Q(i)+Q(i,j);
        end
     %%%%%%%%%%%%%%%%%%%%从首端向末端推电压%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
         Vbus(1) = 66*1.0500;
         for s=branchnum:-1:1
                i=T1(s,2); 
                j=T1(s,3);
                R=T1(s,4);
                X=T1(s,5);
                Vbus(j)=(Vbus(i))^2-2*(P(j)*R+Q(j)*X)+(R*R+X*X)*(P(j)*P(j)+Q(j)*Q(j))/(Vbus(i)*1000)^2;%始端电压计算
                Vbus(j)=sqrt(Vbus(j));
         end
    e=max(abs(Vbus1-Vbus));%收敛条件
    Vbus1=Vbus;
   k=k+1;
end
fprintf('迭代次数k=%d\n',k);
fprintf('电压幅值:\n');
disp(Vbus1);
fprintf('有功分量:\n');
disp(P);
fprintf('无功分量:\n');
disp(Q);
for i=1:14
    A(i)=i;
end
A=A';
plot(A,Vbus1,'linewidth',1.5);
xlabel('节点序号');
ylabel('节点电压幅值');
title('节点电压幅值图');
xlswrite('首端电压.xlsx',Vbus1);
xlswrite('OutPower.xlsx',P);%线路i->j的首端功率
