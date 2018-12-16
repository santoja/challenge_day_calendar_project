#!/bin/bash

#monday
MONDAY=$(date -v-monday -v-4w +'%d/%m/%Y')

echo $MONDAY

