# delphi
一款基于delphi的表格控件，先简单的实现了单元格的合并</br>
已经实现：</br>

<b>单元格合并</b><br>
可以支持单元格的合并，使用方法示例：<br>
<pre>
  miniGrid.MergeCells(1, 1, 1, 1);//以第一列第一行为准，合并1列和1行<br>
  miniGrid.MergeCells(3, 3, 0, 1);//以第三列第三行为准，合并0列和1行<br>
</pre>

<b>单元格自动超链接自动识别</b>
使用示例：<br>
<pre>
  miniGrid.Cells[4,1] := 'http://www.cnblogs.com/5207/';<br>
  miniGrid.Cells[4,2] := '<A href="http://www.cnblogs.com/5207/">Click here</A>';<br>
</pre>
  


