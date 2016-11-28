#!/bin/bash

run=0
passed=0
failed=0

cd $PWD

cat /dev/null > passed.txt
cat /dev/null > failed.txt
cat /dev/null > mail.txt
echo "Running test scripts..."

while read line
  do
    run=$(($run+1))
    cat /dev/null > result.jtl
    cat /dev/null > result.tmp
    cat /dev/null > casename.tmp
    echo $line
    sh /Users/wangrong/VEX/apache-jmeter-3.0/bin/jmeter.sh -n -t $line -l result.jtl | grep -Eo "\/.+jmx" >> casename.tmp
    elapsed=0
    cat result.jtl | while read line2
      do
        awk -F ',' '$8 == "false" {print "Step: "$3"\nReason: "$9}' >> result.tmp
      done
    if [ -s result.tmp ]
      then
        failed=$(($failed+1))
        cat casename.tmp >> failed.txt
        echo Elapsed: `awk -F "," '{elapsed += $2};END {print elapsed}' result.jtl` ms >> failed.txt
        cat result.tmp >> failed.txt
        echo "-----------------------------------------" >> failed.txt
      else
        passed=$(($passed+1))
        cat casename.tmp >> passed.txt
        echo Elapsed: `awk -F "," '{elapsed += $2};END {print elapsed}' result.jtl` ms >> passed.txt
        echo "-----------------------------------------" >> passed.txt
    fi
  done < $1
find /usr/local/automation/logs/ -name 'PostTest*.log' -type f -mmin -30 -print |xargs grep App | awk -F "&quot;" '{print $2 ,$4}' > version
#cat $profile | grep project.name >> mail.txt
#cat $profile | grep project.version >> mail.txt
#cat $profile | grep build.date >> mail.txt
#cat $profile | grep build.revision >> mail.txt
echo "Run ${run} cases." >> mail.txt
echo "${passed} cases passed." >> mail.txt
echo "${failed} cases failed." >> mail.txt
echo "Attached please find the test results." >> mail.txt
#echo "------------------------------------" >> mail.txt
echo "++++++++++++build version+++++++++++++++++" >> mail.txt
find /usr/local/automation/logs/ -name 'PostTest*.log' -type f -mmin -30 -print |xargs grep App | awk -F "&quot;" '{print $2 " :"$4}' | awk 'NR==1,NR==3' >> mail.txt
echo "==========================================" >> mail.txt
find /usr/local/automation/logs/ -name 'PostTest*.log' -type f -mmin -30 -print |xargs grep App | awk -F "&quot;" '{print $2 " :"$4}' | awk 'NR==5,NR==7' >> mail.txt
echo "==========================================" >> mail.txt
find /usr/local/automation/logs/ -name 'PostTest*.log' -type f -mmin -30 -print |xargs grep App | awk -F "&quot;" '{print $2 " :"$4}' | awk 'NR==9,NR==11' >> mail.txt
echo "==========================================" >> mail.txt
find /usr/local/automation/logs/ -name 'PostTest*.log' -type f -mmin -30 -print |xargs grep App | awk -F "&quot;" '{print $2 " :"$4}' | awk 'NR==13,NR==15' >> mail.txt
echo "+++++++++++++++++++++++++++++++++++++++++++" >> mail.txt  
 

while read recp
  do
    mutt -s "Vex Daily Test Result" -a passed.txt failed.txt -- $recp < mail.txt
  done < $2

echo "Mail has been sent."
