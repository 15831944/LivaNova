   M S B A M L   `   `   `  ���� ?  ;SSDL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SSDLSSDL      SSDL.PDFLayerview   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     $localclr-namespace:SSDL     -   ���	#FF5C5C5C  � Title$  PDF Layer Conversion��$	��420��$	��500�� � WindowStyle$ 
ToolWindow=� � WindowStartupLocation$ CenterScreen=� � 	Resources  �� 
TargetType �� Property �� Value �� 
TargetType J� SnapsToDevicePixels	 {�Name
 �� Triggers P� Property P� Value �� 
TargetName N� EnterActions �� 
TargetName �� TargetProperty t� From	 t� To g� Duration g� AutoReverse g� RepeatBehavior��%�  '�       �� " �.���  � �� �	#FFCBECB0��  � ��
 �������  � ��
 �������  � ��$ 1q���  � �� �� " �.R�� $��TemplateRoot�� . .+��� #��z���� $��
PART_Track��$��0q�{� $	 PART_Indicator��$
��Left=�$��0q�._�� $��
Foreground��.+��� $��	Indicator��#��z��� $��	Animation��}�. .+��� $��PART_GlowRect��$��0,0,0,0q�$
��Left=�#��z��� $��Overlay���� #��z��#��z��
 P�  �@�IsIndeterminate . . ��� $ 	Indicator�� � ��
 ���� �� $ PART_GlowRect�� � ��$	 100�� �� .���� .��t� $ PART_GlowRect��$ Margin�$ 	-50,0,0,0X�$ 	400,0,0,0X�$ 0:0:2D� . $ Forever��.���  � RowDefinitions �� $��5 ��� $��30 ��� �� $��30 ��� $
��Auto ��� $��5 � � ColumnDefinitions �� $��5 ��� $��* ��� $��70 ��� $��5 �.+��� $��1��$��1��$��
Layer List��
�������$��Center=�$
��Left=�&   System.Windows.Controls.DataGrid  -   $��dgvLayerlist��
���   �$��1��$��2��$��2��   CanUserAddRows .     CanUserDeleteRows .     AutoGenerateColumns .  ����    Columns 4  .System.Windows.Controls.DataGridTemplateColumn  ,  &System.Windows.Controls.DataGridColumn   Width5  /System.Windows.Controls.DataGridLengthConverter$ 50    Header$ Include��   CellTemplate �� .R��� $��Center=� d� 	IsChecked �� *	Include+  �� Mode$  TwoWay=�! �� UpdateSourceTrigger$! PropertyChanged=�0  *System.Windows.Controls.DataGridTextColumn  $ 
Layer Name��$	 75* "   
IsReadOnly" . 1  +System.Windows.Controls.DataGridBoundColumn#   Binding# �� *	LayerName+$  TwoWay=�  $	 100 $ Contains Data��" .  �� .R��� $��Center=�{�.   �� *EntityAvail+$  TwoWay=��� -   $��BtnGenerate��$��1��$��3��$��25��$��Right=�$��67��$��3q�$��Generate���� -   $��BtnClose��$��2��$��3��$��25��$��3,3,0,3q�$��Close��� -   $��gridProcess��$��4��$��0q�$y�	Collapsed=�$��3�� �� $��30 ��� $��30 ��� $��* �.+��� -   $��	lblHeader��
�������$��0��$��Processing......���� -   $��lblStatusPublisher��
�������$��1��� -   $��pbStatus��$��2��$��5,0,5,0q�$��30���� -	   $��txtPercentage��$��2��
���  ��$��Center=�$��Center=�