#!/bin/bash
rm -rf tests/*.xlsx
for filename in tests/*.xlsm; do
  [ -e "$filename" ] || continue 
  if [[ $filename != *"schedule.xlsm" ]];then
    echo $filename
    $PWD/schedule.py "$PWD/$filename"
    echo -e ''
  fi
done
