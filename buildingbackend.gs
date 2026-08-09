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
