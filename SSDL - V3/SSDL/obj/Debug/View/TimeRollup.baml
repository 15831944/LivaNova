   M S B A M L   `   `   `  ����?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:SSDL.View	SSDL.View  5   	       SSDL.View.Timerollup   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   )localclr-namespace:SSDL.View     5   	   -   5          � 
ResizeMode$  NoResize=�5   /    � WindowStartupLocation$ CenterScreen=�6E   ���	#FF5C5C5C � Title$ Time Rollup��5   	   $	��600��6   $	��400��6*   .��5
      �  � RowDefinitions 5   
   �� 5      $	��100 �6   �� 5      $��* �6   .+�5   
   � $��0��6   .+�5      �� -   5         ��Name$ label��5      $��Enter name of BOM��6#   $
��Left=�6?   $��5,10,0,0q�6Z   
�������$	��Top=�6   �� 5      -   5         ��Name$ textBoxdesc��5      $��23��6+   $��
10,43,10,0q�67    �� TextWrapping$
 Wrap=�6K   $�� ��6_   $	��Top=�6g   6   �� 5      -   5        $ lblRemainingText��5      $�� ��6.   
�������$
��Left=�6L   $��Right=�6g   $��
10,70,10,0q�6�   $	��Top=�6�   ��6�   ��  �� Path$ ActualWidth� �� ElementName$ textBoxdesc��5      � 5   
   $��1��6    5      �� 5      $��1 �6    �� 5      $��* �6    �� 5      $��35 �6    �� 5      $��40 �6    	 � ColumnDefinitions	 5      �� 5      $��5 �6#   �� 5      $��* �6#   �� 5      $��5 �6#   .+�5       �� -   5        $��	TotalMins��5    K   $��Right=�6   $��Center=�60   $��Total Mins :��6\   $��2��6s   $��1��6�   
�������6�   �� 5!      -   5        $��BtnOk��5!      $��3��6"   $��1��6/   $��70��6?   $��Ok��6J   $��Right=�6W   $��5,5,5,5q�6s   6�   &   System.Windows.Controls.DataGrid  5"      -   5        $��dgvTimeRollup��5"      
���   �$��0��6A   $��3��6Q   $��1��6e   
   CanUserAddRows
 .     CanUserDeleteRows .     AutoGenerateColumns .     Columns 5$      4  .System.Windows.Controls.DataGridCheckBoxColumn  5%      ,  &System.Windows.Controls.DataGridColumn   Width5  /System.Windows.Controls.DataGridLengthConverter$ 85 6-   1  +System.Windows.Controls.DataGridBoundColumn   Binding 68   �� *	include+ �� UpdateSourceTrigger$ PropertyChanged=� �� Mode$ TwoWay=�   Header 5&      �� 5'      -   5         ��Name$ ckIncludedbHeader��5'   '   $��
Included ?��6B   6�      	CellStyle 5)      �� 5*      .�5+   "   -	   5        5-      5/      0  *System.Windows.Controls.DataGridTextColumn  50      $ Category��6)   $
 100* 6;      
IsReadOnly .   6H   �� *
Category+6e     51      $
 Time��6)   $	 50* 67    .   6C   �� *Time+6\   53      54      55      57      