   M S B A M L   `   `   `  ����?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SSDLSSDL  5   	   D @DWGViewer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null8*clr-namespace:DWGViewer;assembly=DWGViewer	DWGViewer 5   	       SSDL.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:SSDL     5   	   =viewer*clr-namespace:DWGViewer;assembly=DWGViewer     5   	   -   5        ���	#FF5C5C5C  � Title$  Open Drawing��5
   	   $	��600��6   $	��750��6,    � WindowStartupLocation$ CenterScreen=�68    � 
ResizeMode$ NoResize=�6]   .��5      �  � RowDefinitions 5   
   �� 5      $��5* �6   �� 5      $	��26* �6   �� 5      $
��132* �6   �� 5      $	��37* �6   �� 5      $	��26* �6   �� 5      $	��36* �6   �� 5      $
��224* �6   �� 5      $	��40* �6    � ColumnDefinitions 5   
   �� 5      $	��260 �6   �� 5      .+�5   
   �� $��0��6   $��Look in:��6    $
��Left=�63   
�������$	��Top=�6a   $��1��6y   $��26��6�   $��51��6�   �� 5   
   -   5         ��Name$ BtnRef��5      $��20��6!   
������ $��0��6F   $��Right=�6V   $��1��6�   $��Bottom=�6�   $��1,1,6,1q�6�   .��5      � $��Center=�6   .+�5      �� $!��/SSDL;component/Refresh.ico��6   $��16��6=   $��Bottom=�6I   5      5       R� 5!   
   -   5         R�Name$ foldersItem��5!      $��2��6(   $��0��65   $	��120��6F   $	��Top=�6S   $��10,5,6,0q�6k   6}   �� 5#   
   $��3��6   $��0��6   
�������$��Find:��6@   $
��Left=�6P   $��Center=�6k   $��26��6�   $��0,5,0,6q�6�   $��35��6�   �� 5$   
   -   5         ��Name$ TxtFind��5$      $��0��6#   $��40,2,6,2q�63   $��22��6E   $��3��6Q   $��Center=�6`   �� 5%   
   -   5        $ BtnFind��5%      $
��Find��6"   $��0��61   $��Right=�6A   $��4��6^   $��Center=�6k   $��46��6�   $��22��6�   $��0,2,6,2q�6�   6�   �� 5'   
   $��5��6   $��0��6   
�������$
��Left=�6@   $��Center=�6[   $��Name:��6v   $��0,4,0,6q�6�   $��26��6�   $��44��6�   �� 5(   
   -   5         ��Name$ LstFiles��5(      $	��210��6$   $��10,5,6,0q�61   $��6��6C   $��0��6P   $	��Top=�6a   6�   �� 5)   
   -   5        $ BtnOpen��5)      $��0��6"   $
��Open��62   $��Right=�6A   $��7��6^   $��Center=�6k   $��60��6�   $��25��6�   $��0,1,75,7q�6�   6�   �� 5*   
   -   5        $ BtnClose��5*      $��Close��6#   $��0��63   $��Right=�6C   $��7��6`   $��Center=�6m   $��60��6�   $��25��6�   $��5,1,6,7q�6�   6�   �� 5,   
   -	   5        $��
chkPreview��5,      $��2��6%   $��	Preview :��65   $
��Left=�6I   $��Bottom=�6d   
�������$��1��6�   $��21��6�   $��0,0,0,0q�6�   6�   �� 5-   
   $��1q�6   $��2��6&   $��2��66   $��5��6D   ._�5.      ]� -
   5        $��	dpPreview��5.      $��Stretch=�6)   $��Stretch=�6G   ��. .+�50      7  1System.Windows.Forms.Integration.WindowsFormsHost  -   5        $��host1��50   6   	   
Background
	 �����$��0,0,0,0q�6C   $��Stretch=�6T   $��Stretch=�6r   52      53      54      55      56      