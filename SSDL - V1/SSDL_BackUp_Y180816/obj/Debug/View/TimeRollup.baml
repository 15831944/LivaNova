   M S B A M L   `   `   `  ����?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:SSDL.View	SSDL.View  5   	       SSDL.View.Timerollup   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   )localclr-namespace:SSDL.View     5   	   -   5          � 
ResizeMode$  NoResize=�5   /   ���	#FF5C5C5C � Title$ Time Rollup��5   	   $	��600��6   $	��400��6*   .��5
      �  � RowDefinitions 5   
   �� 5      $��1 �6   �� 5      $��* �6   �� 5      $��35 �6   �� 5      $��40 �6    � ColumnDefinitions 5   
   �� 5      $��5 �6   �� 5      $��* �6   �� 5      $��5 �6   .+�5   
   �� -   5        $��	TotalMins��5   G   $��Right=�6   $��Center=�6,   $��Total Mins :��6X   $��2��6o   $��1��6|   
�������6�   �� 5   
   -   5        $��BtnOk��5      $��3��6   $��1��6+   $��70��6;   $��OK��6F   $��Right=�6S   $��5,5,5,5q�6o   6�   &   System.Windows.Controls.DataGrid  5   
   -   5        $��dgvTimeRollup��5      
���   �$��0��6=   $��3��6M   $��1��6a      CanUserAddRows .     CanUserDeleteRows .     AutoGenerateColumns .     Columns 5      4  .System.Windows.Controls.DataGridCheckBoxColumn  5      ,  &System.Windows.Controls.DataGridColumn   Width5  /System.Windows.Controls.DataGridLengthConverter$ 70 6)   	   Header$	 
Included ?��6�   1  +System.Windows.Controls.DataGridBoundColumn
   Binding
 64   �� *	include+ �� UpdateSourceTrigger$ PropertyChanged=� �� Mode$ TwoWay=�   	CellStyle 5      �� 5       .�5!      -   5        5#      5%      0  *System.Windows.Controls.DataGridTextColumn  5&      $	 Category��6%   $	 75* 67      
IsReadOnly .  
 6C   �� *
Category+6`     5'      $
	 Time��6%   $	 75* 63    .  
 6?   �� *Time+6X   5)      5+      5-      