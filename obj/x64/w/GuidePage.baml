   M S B A M L   `   `   `  ���� G  CAdmin Coffee, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:Admin_CoffeeAdmin_Coffee      Admin_Coffee.GuidePage   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     ,localclr-namespace:Admin_Coffee       N� Title$  	GuidePage��.��� ���#1E1E1E � ColumnDefinitions �� $��7* ��� $��3* �.+��� $��Stretch=�$��Stretch=�.+�&   System.Windows.Controls.DataGrid  -     Name$ dataGrid_guides��$��Stretch=�$��10q�$	��400��   AutoGenerateColumns .     CanUserAddRows .  $��12(�$��Century Gothic)�   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Код��   Width5  /System.Windows.Controls.DataGridLengthConverter$ 1*    
IsReadOnly . 1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 �� 
 �� Path$

 code�  $  ��$ 1*  .  	 �� $
 title�  $  ��$ 1*  .  	 �� $
 title2�  $  ��$ 1*  .  	 �� $
 title3�  $  ��$ 1*  .  	 �� $
 title4�  $  ��$ 1*  .  	 �� $
 title5�  $  ��$ 1*  .  	 �� $
 title6�  $  ��$ 1*  .  	 �� $
 title7�  $  ��$ 1*  .  	 �� $
 title8�  $  ��$ 1*  .  	 �� $
 title9��� $��1��.+��� -    ��Name$ addBtn��$#��Добавить запись��$��Center=�$��10, 15, 10, 0q�$	��Top=�$��35��$	��133��
������ 
�������$��12(�$��Century Gothic)� � Cursor$
 Hand��
��������� -   $ saveBtn��$��Сохранить��$��Center=�$��10, 15, 10, 0q�$	��Top=�$��35��$	��133��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ 	deleteBtn��$��Удалить��$��Center=�$��10, 15, 10, 20q�$	��Top=�$	��133��$��35��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ 
clientsBtn��$��Клиенты��$��Center=�$��1q�$	��Top=�$	��132��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ postBtn��$��Должности��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ smetaBtn��$��
Смета��$��Center=�$��1q�$	��Top=�$	��133��$��20�� J� RenderTransformOrigin$ 1.323,1.559&�
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ menuBtn��$��Меню��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -	   $ salesBtn��$ ��Акции и скидки��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -
   $ suppliersBtn��$��Поставщики��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ furniture_registrBtn��$��Реестр мебели��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ advBtn��$��Реклама��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ staffBtn��$��Сотрудники��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ trade_turnoverBtn��$��Товарооборот��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
��������� -   $ 	branchBtn��$��Филиал��$��Center=�$��1q�$	��Top=�$	��133��$��20��
������ 
�������$��12(�$��Century Gothic)�$
 Hand��
�������