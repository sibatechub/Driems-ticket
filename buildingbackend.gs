function getBuildings() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const ticketSheet = ss.getSheetByName("Master");

  const buildingData = buildingSheet.getDataRange().getValues();
  buildingData.shift();

  const ticketData = ticketSheet.getDataRange().getValues();
  ticketData.shift();

  // Store ticket count for each building
  const ticketCount = {};

  ticketData.forEach(function(row){

    const building = row[8];   // Building Name (Column I)

    if (!building) return;

    ticketCount[building] = (ticketCount[building] || 0) + 1;

  });

const result = buildingData.map(function(row){

    return {

        slNo: row[0],
        building: row[1],
        technician: row[2],
        email: row[3],
        mobile: row[4],
        status: row[5],
        tickets: ticketCount[row[1]] || 0

    };

});

// Highest ticket count first
result.sort(function(a, b){
    return b.tickets - a.tickets;
});

return result;

}
function getBuildingStatistics() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const workloadSheet = ss.getSheetByName("WorkLoad");

  const buildingData = buildingSheet.getDataRange().getValues();
  buildingData.shift();

  const totalBuildings = buildingData.length;

  let activeBuildings = 0;
  let technicians = {};

  buildingData.forEach(function(row){

    if(String(row[5]).toUpperCase() == "YES"){
      activeBuildings++;
    }

    if(row[2]){
      technicians[row[2]] = true;
    }

  });

  const workloadData = workloadSheet.getDataRange().getValues();
  workloadData.shift();

  let overrideActive = 0;

  const today = new Date();
  today.setHours(0,0,0,0);

  workloadData.forEach(function(row){

      const active = String(row[8]).toUpperCase();

      if(active != "YES") return;

      const start = new Date(row[5]);
      const end = new Date(row[6]);

      start.setHours(0,0,0,0);
      end.setHours(23,59,59,999);

      if(today >= start && today <= end){
          overrideActive++;
      }

  });

  return {

      totalBuildings: totalBuildings,
      activeBuildings: activeBuildings,
      overrideActive: overrideActive,
      technicians: Object.keys(technicians).length

  };

}
function getBuildingForEdit(buildingName) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Building Sheet
  const buildingSheet = ss.getSheetByName("Building_Technician_Map");
  const buildingData = buildingSheet.getDataRange().getValues();

  // Users Sheet
  const usersSheet = ss.getSheetByName("Users");
  const usersData = usersSheet.getDataRange().getValues();

  let building = null;

  // Find selected building
  for (let i = 1; i < buildingData.length; i++) {

    if (String(buildingData[i][1]) == String(buildingName)) {

      building = {
        building: buildingData[i][1],
        technician: buildingData[i][2],
        email: buildingData[i][3],
        mobile: buildingData[i][4],
        status: buildingData[i][5]
      };

      break;

    }

  }

  // Active technicians only
  const technicians = [];

  for (let i = 1; i < usersData.length; i++) {

    const role = String(usersData[i][6]).trim().toUpperCase();
    const status = String(usersData[i][7]).trim().toUpperCase();

    if (role == "TECHNICIAN" && status == "YES") {

      technicians.push({

        name: usersData[i][2],
        email: usersData[i][3],
        mobile: usersData[i][4]

      });

    }

  }

  return {

    building: building,
    technicians: technicians

  };

}
function updateBuilding(building){

  const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Building_Technician_Map");

  const data = sheet.getDataRange().getValues();

  for(let i=1;i<data.length;i++){

    if(String(data[i][1]) == String(building.building)){

      // Technician
      sheet.getRange(i+1,3).setValue(building.technician);

      // Email
      sheet.getRange(i+1,4).setValue(building.email);

      // Mobile
      sheet.getRange(i+1,5).setValue(building.mobile);

      // Status
      sheet.getRange(i+1,6).setValue(building.status);

      return true;

    }

  }

  return false;

}
