FasdUAS 1.101.10   ��   ��    k             l     ��  ��    N HA script to write files in conjunction with ExportToHTML MS Excel add-in     � 	 	 � A   s c r i p t   t o   w r i t e   f i l e s   i n   c o n j u n c t i o n   w i t h   E x p o r t T o H T M L   M S   E x c e l   a d d - i n   
  
 l     ��  ��    2 ,Valeriy Kosikhin (vkosikhin@gmail.com), 2016     �   X V a l e r i y   K o s i k h i n   ( v k o s i k h i n @ g m a i l . c o m ) ,   2 0 1 6      l     ��  ��    x rMust be placed in the "/Users/Username/Library/Application Scripts/com.microsoft.Excel/" folder as WrieToFile.scpt     �   � M u s t   b e   p l a c e d   i n   t h e   " / U s e r s / U s e r n a m e / L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / "   f o l d e r   a s   W r i e T o F i l e . s c p t      l     ��������  ��  ��        i         I      �� ���� 0 writetofile WriteToFile   ��  o      ���� 0 paramstring ParamString��  ��    k     2       l     ��������  ��  ��         r      ! " ! I      �� #���� 0 splitstring SplitString #  $ % $ o    ���� 0 paramstring ParamString %  &�� & m     ' ' � ( (  ; ; ;��  ��   " J       ) )  * + * o      ���� 0 
outputfile 
OutputFile +  ,�� , o      ���� 0 outputstring OutputString��      - . - l   ��������  ��  ��   .  / 0 / I   �� 1 2
�� .rdwropenshor       file 1 o    ���� 0 
outputfile 
OutputFile 2 �� 3��
�� 
perm 3 m    ��
�� boovtrue��   0  4 5 4 I   &�� 6 7
�� .rdwrseofnull���     **** 6 o     ���� 0 
outputfile 
OutputFile 7 �� 8��
�� 
set2 8 m   ! "����  ��   5  9 : 9 I  ' 0�� ; <
�� .rdwrwritnull���     **** ; l  ' ( =���� = o   ' (���� 0 outputstring OutputString��  ��   < �� > ?
�� 
refn > o   ) *���� 0 
outputfile 
OutputFile ? �� @��
�� 
as   @ m   + ,��
�� 
utf8��   :  A�� A l  1 1��������  ��  ��  ��     B C B l     ��������  ��  ��   C  D�� D i     E F E I      �� G���� 0 splitstring SplitString G  H I H o      ���� 0 	bigstring 	BigString I  J�� J o      ����  0 fieldseparator FieldSeparator��  ��   F k      K K  L M L l     ��������  ��  ��   M  N O N r      P Q P 1     ��
�� 
txdl Q o      ���� 0 oldtid OldTID O  R S R r     T U T o    ����  0 fieldseparator FieldSeparator U 1    
��
�� 
txdl S  V W V r     X Y X n     Z [ Z 2   ��
�� 
citm [ o    ���� 0 	bigstring 	BigString Y o      ���� 0 	textitems 	TextItems W  \ ] \ r     ^ _ ^ o    ���� 0 oldtid OldTID _ 1    ��
�� 
txdl ]  ` a ` l   ��������  ��  ��   a  b c b L     d d o    ���� 0 	textitems 	TextItems c  e�� e l   ��������  ��  ��  ��  ��       �� f g h��   f ������ 0 writetofile WriteToFile�� 0 splitstring SplitString g �� ���� i j���� 0 writetofile WriteToFile�� �� k��  k  ���� 0 paramstring ParamString��   i �������� 0 paramstring ParamString�� 0 
outputfile 
OutputFile�� 0 outputstring OutputString j  '������������������������ 0 splitstring SplitString
�� 
cobj
�� 
perm
�� .rdwropenshor       file
�� 
set2
�� .rdwrseofnull���     ****
�� 
refn
�� 
as  
�� 
utf8�� 
�� .rdwrwritnull���     ****�� 3*��l+ E[�k/E�Z[�l/E�ZO��el O��jl O����� OP h �� F���� l m���� 0 splitstring SplitString�� �� n��  n  ������ 0 	bigstring 	BigString��  0 fieldseparator FieldSeparator��   l ���������� 0 	bigstring 	BigString��  0 fieldseparator FieldSeparator�� 0 oldtid OldTID�� 0 	textitems 	TextItems m ����
�� 
txdl
�� 
citm�� *�,E�O�*�,FO��-E�O�*�,FO�OPascr  ��ޭ