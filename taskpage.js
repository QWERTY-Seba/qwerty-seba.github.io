Office.actions.associate('PASTECLIPBOARD', function () {
    console.log("DEBUGINS")
});


Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {

    }
});



