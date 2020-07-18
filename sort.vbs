Dim result(9)'-------配列
Dim temp'------------出力用
Randomize'-----------ランダムシステム準備

'------ランダムな数値を配列にセット-----------------
for i = 0 to 9
  result(i) = i
next
for i = 0 to 9
  num = Int(Rnd() * 9)
  w = result(i)
  result(i) = result(num)
  result(num) = w
next

'------配列のデータを出力用にセット-----------------
for i = 0 to 9
  temp =  temp & result(i) & ","
next
'WScript.Echo temp

'------並べ替え（Bubbleソート）---------------------
for i = 0 to 8
  for j = 0 to 8-i
    if result(j) > result(j+1) then
        w = result(j)
        result(j) = result(j+1)
        result(j+1) = w
    end if
  next
next

'------整列後の配列のデータを出力用にセット---------
temp = temp + "   "
for i = 0 to 9
  temp =  temp & result(i) & ","
next
WScript.Echo temp