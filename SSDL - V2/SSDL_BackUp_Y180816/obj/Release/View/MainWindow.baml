   M S B A M L   `   `   `  ���� ?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SSDLSSDL  D @DWGViewer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null8*clr-namespace:DWGViewer;assembly=DWGViewer	DWGViewer     SSDL.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     $localclr-namespace:SSDL     =viewer*clr-namespace:DWGViewer;assembly=DWGViewer     -   ���	#FF5C5C5C  � Title$  Open Drawing��$	��600��$	��750�� � WindowStartupLocation$ CenterScreen=� � 
ResizeMode$ NoResize=�.���  � RowDefinitions �� $��5* ��� $	��26* ��� $
��132* ��� $	��37* ��� $	��26* ��� $	��36* ��� $
��224* ��� $	��40* � � ColumnDefinitions �� $	��260 ��� .+��� $��0��$��Look in:��$
��Left=�
�������$	��Top=�$��1��$��26��$��51���� -    ��Name$ BtnRef��$��20��
������ $��0��$��Right=�$��1��$��Bottom=�$��1,1,6,1q�.��� $��Center=�.+��� $!��/SSDL;component/Refresh.ico��$��16��$��Bottom=�R� -    R�Name$ foldersItem��$��2��$��0��$	��120��$	��Top=�$��10,5,6,0q��� $��3��$��0��
�������$��Find:��$
��Left=�$��Center=�$��26��$��0,5,0,6q�$��35���� -    ��Name$ TxtFind��$��0��$��40,2,6,2q�$��22��$��3��$��Center=��� -   $ BtnFind��$
��Find��$��0��$��Right=�$��4��$��Center=�$��46��$��22��$��0,2,6,2q��� $��5��$��0��
�������$
��Left=�$��Center=�$��Name:��$��0,4,0,6q�$��26��$��44���� -    ��Name$ LstFiles��$	��210��$��10,5,6,0q�$��6��$��0��$	��Top=��� -   $ BtnOpen��$��0��$
��Open��$��Right=�$��7��$��Center=�$��60��$��25��$��0,1,75,7q��� -   $ BtnClose��$��Close��$��0��$��Right=�$��7��$��Center=�$��60��$��25��$��5,1,6,7q��� -	   $��
chkPreview��$��2��$��	Preview :��$
��Left=�$��Bottom=�
�������$��1��$��21��$��0,0,0,0q��� $��1q�$��2��$��2��$��5��._�]� -
   $��	dpPreview��$��Stretch=�$��Stretch=���. .+�7  1System.Windows.Forms.Integration.WindowsFormsHost  -   $��host1��	   
Background
	 �����$��0,0,0,0q�$��Stretch=�$��Stretch=�