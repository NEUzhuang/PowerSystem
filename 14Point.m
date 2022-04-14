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
%% ����
%%%%%%%%%%%%%%%֧·�������򣬸�����֧��ͬʱ���м���
while ~isempty(Branch1)%%%%T1��֧·����
    m=1;
    [s,row]=size(Branch1);
        while s>0
            t=find(Branch1(:,2)== Branch1(s,3));%�ж��Ƿ����ӽڵ�
                if isempty(t) %�ӽڵ�
                    T1(n,:)= Branch1(s,:);%�ӽڵ�ϵͳĩ�����׶˽��У�
                    n=n+1;
                else
                    T2(m,:)= Branch1(s,:);%�����ӽڵ�
                     m=m+1;
                end
               s=s-1;
        end
        Branch1=T2;
        T2=[];
end
%% 
%%%%%%%%%%%%%%%%%%%%%%%
while e>1.0e-07%����
    %%%%%%%%%%%%%%%%%%%%%%%%��ĩ�����׶��ƹ���%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    P=zeros(busnum,1);%%��Ź���
    Q=zeros(busnum,1);
        for s=1:branchnum
                i=T1(s,2);
                j=T1(s,3);
                R=T1(s,4);
                X=T1(s,5);
                Pload=Bus(j,2);%�й�����
                Qload=Bus(j,3);%�޹�����
                II=((Pload+P(j))^2+(Qload+Q(j))^2)/(Vbus(j)^2*1000);
                Ploss(i,j)=II*R;%֧·�й����
                Qloss(i,j)=II*X;%֧·�޹����
                P(i,j)=Pload+Ploss(i,j)+P(j);
                Q(i,j)=Qload+Qloss(i,j)+Q(j);
                P(i)=P(i)+P(i,j);
                Q(i)=Q(i)+Q(i,j);
        end
     %%%%%%%%%%%%%%%%%%%%���׶���ĩ���Ƶ�ѹ%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
         Vbus(1) = 66*1.0500;
         for s=branchnum:-1:1
                i=T1(s,2); 
                j=T1(s,3);
                R=T1(s,4);
                X=T1(s,5);
                Vbus(j)=(Vbus(i))^2-2*(P(j)*R+Q(j)*X)+(R*R+X*X)*(P(j)*P(j)+Q(j)*Q(j))/(Vbus(i)*1000)^2;%ʼ�˵�ѹ����
                Vbus(j)=sqrt(Vbus(j));
         end
    e=max(abs(Vbus1-Vbus));%��������
    Vbus1=Vbus;
   k=k+1;
end
fprintf('��������k=%d\n',k);
fprintf('��ѹ��ֵ:\n');
disp(Vbus1);
fprintf('�й�����:\n');
disp(P);
fprintf('�޹�����:\n');
disp(Q);
for i=1:14
    A(i)=i;
end
A=A';
plot(A,Vbus1,'linewidth',1.5);
xlabel('�ڵ����');
ylabel('�ڵ��ѹ��ֵ');
title('�ڵ��ѹ��ֵͼ');
xlswrite('�׶˵�ѹ.xlsx',Vbus1);
xlswrite('OutPower.xlsx',P);%��·i->j���׶˹���
