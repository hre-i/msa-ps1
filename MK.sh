#!/bin/bash

for file in $*
do
    cp -i $file $file.bak

    cat lib/INIT.ps1.in   > $file.new
    sed -e '/^####:INIT:START/,/####:INIT:END/d' \
        -e '/^####:FIN:START/,/^####:FIN:END/d' \
        $file             >> $file.new
    cat lib/FINISH.ps1.in >> $file.new
    mv $file.new $file
done
