#!/bin/sh
#
# An example hook script to verify what is about to be committed.
# Called by "git commit" with no arguments.  The hook should
# exit with non-zero status after issuing an appropriate message if
# it wants to stop the commit.
#
# To enable this hook, rename this file to "pre-commit".


DecompressExcel(){

# File path
properties_file="./src/main/resources/excelversioncontrol.properties"

# Keys in Property File
excelname="excel_file_path"
versioncontrolType="version_control_type"
allExcel="version_control_all_excel"

# Variable to hold the Property Value
prop_value=""
getProperty()
{
prop_key=$1
prop_value=`cat ${properties_file} | grep ${prop_key} | cut -d'=' -f2`
}

#Get the property value
getProperty ${excelname}
excelname=${prop_value}
getProperty ${versioncontrolType}
versioncontrol=${prop_value}
getProperty ${allExcel}
allExcel=${prop_value}


echo $excelname
echo $allExcel


# Get the list of files staged for commit and check if it an excel
filesList=($(git diff --name-only --cached))

if [[ $allExcel == "no" ]];then
IFS=','
read -a excelList <<<"$excelname"
fi

# Get the name of the excel files and pass a sourceFile Parameter to ExcelCompression method
for i in ${!filesList[@]}; do  
    if ([[ $allExcel == "yes" ]] && [[ ${filesList[$i]} =~ \.xlsx$ ]]) || ([[ $allExcel == "no" ]] && [[ ${excelList[*]} =~ ${filesList[$i]} ]]); then
	if [[ $versioncontrol == "xml" ]];then
  		sourceFile=./${filesList[$i]}
		destFolder=${sourceFile/.xlsx/DecompressedFolder}
		method="Decompression"
		# Compile and run the ExcelDecompression
		java -cp ~/Downloads/testversioncontrol/excelversion-ngtp-trunk/src/test/java/tempjar/excel-version-control-1.0.jar excelHelper.ExcelHelper $method $sourceFile $destFolder
    git add $destFolder
 	elif [[ $versioncontrol == "csv" ]];then
		sourceFile=./${filesList[$i]}
		destFolder=${sourceFile/.xlsx/Converted.csv}
		method="ConvertToCSV"

		# Compile and run the ExcelToCSVConvertion
		java -cp ~/Downloads/testversioncontrol/excelversion-ngtp-trunk/src/test/java/tempjar/excel-version-control-1.0.jar excelHelper.ExcelHelper $method $sourceFile $destFolder
	  #git add destFolder

	fi
    fi
done
}


DecompressExcel


