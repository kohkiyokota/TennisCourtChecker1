#!/bin/bash

find /tmp -name ".com.google.Chrome*" -amin +10 -exec rm -r {} \;