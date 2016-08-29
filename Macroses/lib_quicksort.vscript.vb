Sub QuickSort(vec,loBound,hiBound)
  Dim pivot,loSwap,hiSwap,temp

  '== one or 0 item to sort
  if len(vec)<=1 then
    exit sub
  end if

  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(loBound) > vec(hiBound) then
      temp=vec(loBound)
      vec(loBound) = vec(hiBound)
      vec(hiBound) = temp
    End If
  End If

  '== Three or more items to sort
  pivot = vec(int((loBound + hiBound) / 2))
  vec(int((loBound + hiBound) / 2)) = vec(loBound)
  vec(loBound) = pivot
  loSwap = loBound + 1
  hiSwap = hiBound
  
  do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(loSwap) <= pivot
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(hiSwap) > pivot
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then
      temp = vec(loSwap)
      vec(loSwap) = vec(hiSwap)
      vec(hiSwap) = temp
    End If
  loop while loSwap < hiSwap
  
  vec(loBound) = vec(hiSwap)
  vec(hiSwap) = pivot
  
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then QuickSort(vec,loBound,hiSwap-1)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then QuickSort(vec,hiSwap+1,hiBound)

End Sub  'QuickSort

function ArraySort (UnorderedArray)
  if len(UnorderedArray)<>0 then
    QuickSort UnorderedArray,Lbound(UnorderedArray), Ubound(UnorderedArray)
  end if
  ArraySort = UnorderedArray
end function