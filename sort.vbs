Dim result(9)'-------�z��
Dim temp'------------�o�͗p
Randomize'-----------�����_���V�X�e������

'------�����_���Ȑ��l��z��ɃZ�b�g-----------------
for i = 0 to 9
  result(i) = i
next
for i = 0 to 9
  num = Int(Rnd() * 9)
  w = result(i)
  result(i) = result(num)
  result(num) = w
next

'------�z��̃f�[�^���o�͗p�ɃZ�b�g-----------------
for i = 0 to 9
  temp =  temp & result(i) & ","
next
'WScript.Echo temp

'------���בւ��iBubble�\�[�g�j---------------------
for i = 0 to 8
  for j = 0 to 8-i
    if result(j) > result(j+1) then
        w = result(j)
        result(j) = result(j+1)
        result(j+1) = w
    end if
  next
next

'------�����̔z��̃f�[�^���o�͗p�ɃZ�b�g---------
temp = temp + "   "
for i = 0 to 9
  temp =  temp & result(i) & ","
next
WScript.Echo temp