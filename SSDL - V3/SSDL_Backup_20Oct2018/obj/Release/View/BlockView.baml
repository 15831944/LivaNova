   M S B A M L   `   `   `  ���� T  PNumericUpDown, Version=0.1.6.1, Culture=neutral, PublicKeyToken=74df933eb4e7b0ca:-clr-namespace:eisiWare;assembly=NumericUpDowneisiWare  ? ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SSDLSSDL    SSDL.BlockView   :eisiWare-clr-namespace:eisiWare;assembly=NumericUpDown  R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     $localclr-namespace:SSDL     -   ���	#FF5C5C5C  � Title$  Symbol Library��$	��600��$	��750�� � WindowStartupLocation$ CenterScreen=� � 
ResizeMode$ NoResize=�.���  � RowDefinitions �� $��9* ��� $	��36* ��� $
��111* ��� $	��24* ��� $	��35* ��� $	��34* ��� $	��27* ��� $	��87* ��� $	��87* ��� $	��87* ��� $	��35* � � ColumnDefinitions �� $	��186 ��� $��73 ��� .+��� $��0��$��	Look in :��$
��Left=�$��Center=�
�������$��1��$��26��$	��0,5q�$��55���� -    ��Name$ BtnRef��$��20��
������ $��1��$��Right=�$��1��$��Bottom=�$��0,0,10,1q�$��20��.��� $��Center=�.+��� $!��/SSDL;component/Refresh.ico��R� -    R�Name$ foldersItem��$��2��$��0��$��2��$
��Left=�$��93��$	��240��$	��Top=�$��10,0,0,0q��� $��3��$��0��$��Find :��$
��Left=�
�������$	��Top=�$��24��$��39���� -    ��Name$ TxtFind��$��0��$
��Left=�$��22��$��4��$��Center=�$	��240��$��10,7,0,6q�$��2���� -   $ BtnFind��$
��Find��$��1��$��Center=�$��5��$��Center=�$��57��$��22��$	��8,6q��� $��6��$��0��$
��Left=�
�������$	��Top=�$��Name:��$��26��$��44���� -    ��Name$ LstFiles��$
��Left=�$	��246��$��7��$��3��$��0��$��2��$	��Top=�$	��240��$��10,5,0,0q��� -   $ BtnOpen��$��0��$��Insert��$��Right=�$��10��$��Center=�$��60��$��25��$��0,8,5,7q��� -   $ BtnClose��$��Close��$��1��$
��Left=�$��10��$��Center=�$��60��$��25��$��5,8,0,7q��� -	   $��
chkPreview��$��2��$��	Preview :��$
��Left=�$��Center=�
�������$��1��$��27��$��0,10,0,0q��� -
   $��chkclosewindow��$��2��$��Close After Insertion��$��Right=�$��Center=�
�������$��1��$��27��$��	0,10,10,0q��� $��1q�$��2��$��2��$��7��._�7  1System.Windows.Forms.Integration.WindowsFormsHost  -   $��host1��	   
Background
	 �����$��0,0,0,0q��� -   
 ��Name$
 label��$��	Part No :��$��2��
�������$
��Left=�$��9��$	��Top=�$��26��$��56���� -   $ 	TxtPartno��$��2��$
��Left=�$��23��$��63,0,0,0q�$��9�� �� TextWrapping$ NoWrap=�$	��Top=�$	��120��{�.  �� -   $
 label1��$��Qty :��
�������$��2��$
��Left=�$��0,30,0,0q�$��9��$	��Top=�$��26��$��35���� -   $
 label2��$��	Item No :��
�������$��2��$
��Left=�$��0,60,0,0q�$��9��$	��Top=�$��26��$��59��$y�Hidden=��� -   $ 	TxtItemno��$��2��$
��Left=�$��23��$��	63,60,0,0q�$��9��$ NoWrap=�$��1��$	��Top=�$y�Hidden=�$	��120��   eisiWare.NumericUpDown  -     Name$ Txtquantity��$��2��$
��Left=�$��23��$��	63,30,0,0q�$��9��$	��Top=�   Value$ 1Q�$��68���� -   $
 lbllen��$y�Hidden=�$��Len :��
�������$��2��$
��Left=�$��0,60,0,0q�$��9��$	��Top=�$��26��$��59���� -   $ TxtLen��$y�Hidden=�$��2��$
��Left=�$��23��$��	63,60,0,0q�$��9��$ NoWrap=�$�� ��$	��Top=�$	��120��