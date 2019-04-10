#!/bin/bash
rm -rf *.xlsx
for filename in tests/*.xlsm; do
  [ -e "$filename" ] || continue 
  if [[ $filename != *"schedule.xlsm" ]];then
    $PWD/tournament.py "$PWD/$filename" > /dev/null
  fi
done
