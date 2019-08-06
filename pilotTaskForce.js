/*!
* JavaScript application to query a list from a custom search and append the data to an HTML table
* @Reference:
* @Require: jQuery, tablesorter
* @Usage: User inputs data into a SPO form and the data will be displayed (after a custom search) in a formatted table
* @version:
*   1.00 Initial rollout
*/

$(document).on(GIP.Var.eventTriggers.userProfile, function (evt, data) {
    // Clears out the targeted ID in case of subsequent searches
    $("#errorMessage").html("");

    // Set name attributes of three different search criteria
    $("#participantLastName").attr("name", "searchCriteria");
    $("#radioBranchNumber").attr("name", "searchCriteria");
    $("#radioOPSPresidentED").attr("name", "searchCriteria");

    // Controls what happens when the Submit button is clicked
    $("#submitButton").on("click", function (e) {
        // To prevent the page from refreshing itself
        e.preventDefault();
        // Clears out the targeted ID in case of subsequent searches
        $("#errorMessage").html("");

        // Storing user's input in the radio/text fields into variables
        var participantLastName = $("#participantLastName").is(':checked');
        var radioBranchNumber = $("#radioBranchNumber").is(':checked');
        var radioOPSPresidentED = $("#radioOPSPresidentED").is(':checked');
        var fieldSearchQuery = $("#fieldSearchQuery").val();

        // Only generate table if a radio button is selected AND a search query is entered in the form
        if ((participantLastName == true || radioBranchNumber == true || radioOPSPresidentED == true) && fieldSearchQuery !== "") {

            // Storing which radio selection the user picked into a variable
            var radioSelected = (participantLastName == true || radioBranchNumber == true || radioOPSPresidentED == true);

            // Switch statement to change the selected radio choice into a string that is used for the custom AJAX GET request
            switch (radioSelected) {
                case participantLastName:
                    radioSelected = "lastName";
                    break;
                case radioBranchNumber:
                    radioSelected = "branchNumber";
                    break;
                case radioOPSPresidentED:
                    radioSelected = "opsPresidentED";
                    break;
                default:
                    break;
            }

            // Creating a function to grab data and append results into an HTML table with a throw statement to create custom errors
            function getTable() {

                // Call to the Pilot Task Force Participants List
                $.ajax({
                    type: "GET",
                    // OData to order "Pilot Status" field to show "Active" on top + return more than 100 items if needed + dynamically filter based off the user's radio input/search query combination
                    url: "/sites/teams/pilot-taskforce-site/_api/web/lists/getbytitle('pilotTaskForceParticipants')/Items?$orderby=pilotStatus asc&$top=5000&$filter=" + radioSelected + " eq '" + fieldSearchQuery + "'",
                    contentType: "application/json;odata=verbose",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose"
                    },
                    success: function (result) {
                        var data = result.d.results;
                        // Clear HTML table and clear cache in case of subsequent ajax calls
                        $("#pilotTaskForce-table").html("");
                        $("#noResults").html("");
                        $("#pilotTaskForceTable").trigger("update");
                        // Only generate HTML table if there are results
                        if (data.length > 0) {
                            // Loop through the data and append information in correct spots
                            for (var i = 0; i < data.length; i++) {
                                var tRowData = '<tr><td>' + data[i].firstName + '</td><td>' + data[i].lastName + '</td><td>' + data[i].LOB + '</td><td>' + data[i].branchName + '</td><td>' + data[i].pilotTaskForceName + '</td><td>' + data[i].pilotStatus + '</td></tr>';
                                // Appending table row data to the targeted ID
                                $('#pilotTaskForce-table').append(tRowData);
                            }
                            // Table Sorter
                            $("#pilotTaskForceTable").tablesorter({
                                theme: 'red',
                                // Initialize zebra striping of the table
                                widgets: ["zebra", "filter"], 
                                widgetOptions: {
                                    zebra: ["normal-row", "alt-row"],
                                    filter_functions: {
                                        '.z': true
                                    },
                                }
                            });
                        } else {
                            // If no results are returned in the search then inform the user
                            $("#noResults").html("There were no results.");
                            return;
                        }
                    },
                    error: function (error) {
                        throw new Error("Unsuccessful AJAX call");
                    }
                });
            }
        } else {
            // Let the user know that both a "Search Field" and a "Search Term" must be selected
            $("#errorMessage").html("Please Select a \"Search Field\" and Enter in a \"Search Term\".")
            return false;
        }

        // Calling getTable to initialize the AJAX call
        getTable();
        // Show modal if everything is properly filled out and contains results
        $("#taskForceModal").modal("show");
    });
});
