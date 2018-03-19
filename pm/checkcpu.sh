#!/bin/bash

for filename in *
do
echo $filename
sed -n 68508p $filename
done
