function main(workbook: ExcelScript.Workbook) {
    function main(workbook: ExcelScript.Workbook) {
    class UsersClass {
        Id: string;
        UserGroup: string;
        UserName: string;
        Items: string[] = [];
        Fines: string[] = [];
        Total: number;
    }
    //Set the first row to be the 
    let circFineData = workbook.getWorksheet("CircFineData");
    let itemRecordData = workbook.getWorksheet("ItemRecord");
    let userRecordData = workbook.getWorksheet("UserRecord");
    // Run through the item list and then add the values if they are in the list
    let i = 1;
    let userIndex = 0;
    let userList: UsersClass[] = [];
    // Create list of items and issues
    while (circFineData.getCell(i, 0).getValue() != null && circFineData.getCell(i, 0).getValue() !== '') {
        if (circFineData.getCell(i, 0).getValue() !== '-') {
            let idString = circFineData.getCell(i, 0).getValue().toString();
            // Check if list needs to be initialized
            if (userList.length == 0) {
                userList.push(new UsersClass());
                userList[0].Id = idString;
                //Find the UserGroup from UserRecord
                let j = 1;
                while (userRecordData.getCell(j, 0).getValue() != null && userRecordData.getCell(j, 0).getValue() !== '') {
                    if (idString === userRecordData.getCell(i, 0).getValue().toString()) {
                        userList[0].UserGroup = userRecordData.getCell(i, 6).getValue().toString();
                        userList[0].UserName = userRecordData.getCell(i, 9).getValue().toString();
                        break;
                    } else {
                        j++;
                    }
                }
                userList[0].Items.push(circFineData.getCell(i, 1).getValue().toString());
                userList[0].Fines.push(circFineData.getCell(i, 7).getValue().toString());
                userIndex++;
            } else {
                // Check if user is already in list
                let isInUserList = false;
                userList.forEach(userRecord => {
                    if (userRecord.Id === idString) {
                        isInUserList = true;
                        userRecord.Items.push(circFineData.getCell(i, 1).getValue().toString());
                        userRecord.Fines.push(circFineData.getCell(i, 7).getValue().toString());
                    }
                });
                if (!isInUserList) {
                    // Append missing user
                    userList.push(new UsersClass());
                    userList[userIndex].Id = idString;
                    //Find the UserGroup from UserRecord
                    let j = 1;
                    while (userRecordData.getCell(j, 0).getValue() != null && userRecordData.getCell(j, 0).getValue() !== '') {
                        if (idString === userRecordData.getCell(j, 0).getValue().toString()) {
                            userList[userIndex].UserGroup = userRecordData.getCell(i, 6).getValue().toString();
                            userList[userIndex].UserName = userRecordData.getCell(i, 9).getValue().toString();
                            break;
                        } else {
                            j++;
                        }
                    }
                    userList[userIndex].Items.push(circFineData.getCell(i, 1).getValue().toString());
                    userList[userIndex].Fines.push(circFineData.getCell(i, 7).getValue().toString());
                    userIndex++;
                }
            }
        }
        i++;
    }
    // Seperate Users
    let listOfForigvenUsers: UsersClass[] = [];
    let listOfSendUsers: UsersClass[] = [];
    userList.forEach(userRecord => {
        // Check if user if faculty and should be waved
        if (userRecord.UserGroup === 'Faculty') {
            listOfForigvenUsers.push(userRecord);
        } else {
            // Get running total and store in class
            userRecord.Total = 0;
            userRecord.Fines.forEach(fine => {
                let fineArray = fine.split(' ');
                let indexOfFine = fineArray.indexOf('Amount:') + 1;
                userRecord.Total += Number(fineArray[indexOfFine].substring(0, fineArray[indexOfFine].length - 1));
            });
            if (userRecord.Total >= 10) {
                listOfSendUsers.push(userRecord);
            } else {
                listOfForigvenUsers.push(userRecord);
            }
        }
    });
    //Create new list for users
    workbook.addWorksheet('ForgivenUsers');
    workbook.addWorksheet('FinesToSend');
    let forgivenWorksheet = workbook.getWorksheet("ForgivenUsers");
    let finesWorksheet = workbook.getWorksheet("FinesToSend");
    for (let k = 0; k < listOfForigvenUsers.length; k++) {
        forgivenWorksheet.getCell(k, 0).setValue(listOfForigvenUsers[k].Id);
    }
    for (let k = 0; k < listOfSendUsers.length; k++) {
        finesWorksheet.getCell(k, 0).setValue(listOfSendUsers[k].Id);
        finesWorksheet.getCell(k, 1).setValue(listOfSendUsers[k].UserName);
        finesWorksheet.getCell(k, 2).setValue(listOfSendUsers[k].UserGroup);
        finesWorksheet.getCell(k, 3).setValue(listOfSendUsers[k].Total);
        // Get Strings to add items to list and reason
        let stringOfItems = '';
        // Use a loop to keep index instead of having to get it muliptle times in the formula to cominbe reason fee and item info
        for (let l = 0; l < listOfSendUsers[k].Items.length; l++) {
            // Prep fine string for use
            let fineArray = listOfSendUsers[k].Fines[l].split(' ');
            let indexOfFineType = fineArray.indexOf('Fee/Fine type:') + 1;
            let indexOfFineAmount = fineArray.indexOf('Amount:') + 1;
            stringOfItems += " Item: ";
            // Get Title of item
            let titleIndex = 0;
            let m = 1;
            while (itemRecordData.getCell(m, 0).getValue() != null && itemRecordData.getCell(m, 0).getValue() !== '') {
                if (itemRecordData.getCell(m, 7).getValue().toString() === listOfSendUsers[k].Items[l].toString()) {
                    titleIndex = m;
                    break;
                } else {
                    m++;
                }
            }
            stringOfItems += itemRecordData.getCell(titleIndex, 1).getValue();
            stringOfItems += " Barcode: ";
            stringOfItems += listOfSendUsers[k].Items[l].toString();
            stringOfItems += " Fine Type: ";
            stringOfItems += fineArray[indexOfFineType].toString();
            stringOfItems += " Amount: ";
            stringOfItems += fineArray[indexOfFineAmount].toString();
        }
        finesWorksheet.getCell(k, 4).setValue(stringOfItems);
    }
}

}
