   M S B A M L   `   `   `  ����?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SSDLSSDL  5   	       SSDL.PDFLayerview   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:SSDL     5   	   -   5          � 
ResizeMode$  NoResize=�5   *   ���	#FF5C5C5C � Title$ PDF Layer Conversion��5   	   $	��550��6&   $	��500��63    � WindowStartupLocation$ CenterScreen=�6@    � 	Resources 5	       �� 
TargetType �� Property �� Value �� 
TargetType J� SnapsToDevicePixels	 {�Name
 �� Triggers P� Property P� Value �� 
TargetName N� EnterActions �� 
TargetName �� TargetProperty t� From	 t� To g� Duration g� AutoReverse g� RepeatBehavior��%�  '�       �� 5
   
   " �6-   .�5      ��  � �� �	#FFCBECB06+   �� 5       � ��
 �����6+   �� 5       � ��
 �����6@   �� 5       � ��$ 1q�60   �� 5       � �� 5      �� 5      " �6&   .R�5      � $��TemplateRoot��6    . .+�5      �� #��z��6(   �� 5      $��
PART_Track��6(   $��0q�6:   {� 5      $	 PART_Indicator��6(   $
��Left=�6@   $��0q�6[   ._�5   "   � $��
Foreground��6'   .+�5   &   �� $��	Indicator��6T   #��z��60   6T   � 5   &   $��	Animation��6+   }�. .+�5   *   �� $��PART_GlowRect��61   $��0,0,0,0q�6G   $
��Left=�6X   #��z��6s   5   '   � 5   &   $��Overlay��6+   5   '   5   #   5      �� 5       #��z��6%   #��z��5!      5"      
 5#      P� 5$       �@�IsIndeterminate . . �5%   "   �� $ 	Indicator��6)    � ��
 ���� 6P   �� 5&   "   $ PART_GlowRect��6)    � ��$	 100��6U    5'   "   �� 5(   &   .��5)   *   �� .��5*   .   t� $ PART_GlowRect��5+   )   $ Margin�5,   )   $ 	-50,0,0,0X�5-   )   $ 	400,0,0,0X�6;   $ 0:0:2D�6J    . $ Forever��5.   <   5/   +   50   '   52      54      56      57      58      .��59      �  � RowDefinitions 5:   
   �� 5;      $��5 �6   �� 5<      $��30 �6   �� 5=      �� 5>      $��30 �6   �� 5?      $
��Auto �6   6,   �� 5@      $��5 �6    � ColumnDefinitions 5B   
   �� 5C      $��5 �6   �� 5D      $��* �6   �� 5E      $��70 �6   �� 5F      $��5 �6   .+�5I   
   �� $��1��6   $��1��6    $��
Layer List��6-   
�������$��Center=�6U   $
��Left=�6p   &   System.Windows.Controls.DataGrid  5L   
   -   5        $��dgvLayerlist��5L      
���   �$��1��6<   $��2��6L   $��2��6`      CanUserAddRows .     CanUserDeleteRows .     AutoGenerateColumns .  ��5M   i   ��    Columns 5N      4  .System.Windows.Controls.DataGridTemplateColumn  5P      ,  &System.Windows.Controls.DataGridColumn   Width5  /System.Windows.Controls.DataGridLengthConverter$ 50 6)      Header$ Include��64      CellTemplate 5Q      �� 5R      .R�5S      �� $��Center=�6'    d� 	IsChecked 6D   �� *	Include+  �� Mode$  TwoWay=�! �� UpdateSourceTrigger$! PropertyChanged=�6�   5T      5V      0  *System.Windows.Controls.DataGridTextColumn  5W      $ 
Layer Name��6%   $	 75* 69   "   
IsReadOnly" . 1  +System.Windows.Controls.DataGridBoundColumn#   Binding# 6E   �� *	LayerName+$  TwoWay=�6�     5X      $	 100 6)   $ Contains Data��65   $   
Visibility$$ 	Collapsed=�6L   " .  5Y      �� 5Z      .R�5[      �� $��Center=�6'   {�.   6D   �� *EntityAvail+$  TwoWay=�6�   5\      5^      5`      �� 5b   
   -   5        $��BtnGenerate��5b      $��1��6$   $��3��64   $��25��6A   $��Right=�6N   $��67��6j   $��3q�6u   $��Generate��6�   6�   �� 5c   
   -   5        $��BtnClose��5c      $��2��6!   $��3��61   $��25��6>   $��3,3,0,3q�6J   $��Close��6[   6k   � 5d   
   -   5        $��gridProcess��5d   '   $��4��6   $��0q�6   $y�	Collapsed=�6:   $��3��6Q    5e      �� 5f      $��30 �6    6.   �� 5g      $��30 �6    6.   �� 5h      $��* �6    6-   .+�5j      �� -   5        $��	lblHeader��5j      
�������$��0��68   $��Processing......��6E   6b   �� 5k      -   5        $��lblStatusPublisher��5k      
�������$��1��6A   6P   � 5l      -   5        $��pbStatus��5l   8   $��2��6   $��5,0,5,0q�6'   $��30��6H   �� 5m      -	   5        $��txtPercentage��5m   7   $��2��6   
���  ��$��Center=�6L   $��Center=�6j   5n      5o      5p      