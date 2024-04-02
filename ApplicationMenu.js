function fetchAndPopulateDropdown() {
    $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('ApplicationMenuItems')/items",
        type: "GET",
        headers: {
            "Accept": "application/json;odata=verbose"
        },
        success: function (data) {
            var items = data.d.results;
            var dropdown = $('#drpdwn');
            for (var i = 0; i < items.length; i++) {
                var title = items[i].Title;
                var url = items[i].Url;
                var option = $('<option>').val(url).text(title);
                dropdown.append(option);
            }
        },
        error: function (error) {
            console.log(JSON.stringify(error));
        }
    });
}

$(document).ready(function () {
    fetchAndPopulateDropdown();
});
