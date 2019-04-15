#!/bin/bash
rm -rf tests/*.xlsx
for filename in tests/*.xlsm; do
  [ -e "$filename" ] || continue 
  if [[ $filename != *"schedule.xlsm" ]];then
    $PWD/scheduler.py "$PWD/$filename"
  fi
done
