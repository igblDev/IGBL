<!DOCTYPE html>
<html lang="en">
<head>
  <title>Bootstrap Example</title>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
	
<style>
table {
    border-collapse:collapse;
}
table tr td {
    border: 1px solid red;
    padding:2px 15px 2px 15px;
    width:50px;
}
#tabs ul li.drophover {
    color:green;
}
}
</style>
</head>
<body>
<script>
$("#tabs").tabs();

$("tbody").sortable({
    items: "> tr:not(:first)",
    appendTo: "parent",
    helper: "clone"
}).disableSelection();

$("#tabs ul li a").droppable({
    hoverClass: "drophover",
    tolerance: "pointer",
    drop: function(e, ui) {
        var tabdiv = $(this).attr("href");
        $(tabdiv + " table tr:last").after("<tr>" + ui.draggable.html() + "</tr>");
        ui.draggable.remove();
    }
});

</script>
<div id="tabs">
    <ul>
        <li><a href="#current"><span>Current</span></a>

        </li>
        <li><a href="#archive"><span>Archive</span></a>

        </li>
        <li><a href="#future"><span>Future</span></a>

        </li>
    </ul>
    <div id="current">
        <div id="table1">
            <table>
                <tbody>
                    <tr>
                        <td>COL0</td>
                        <td>COL1</td>
                        <td>COL2</td>
                    </tr>
                    <tr>
                        <td>c00</td>
                        <td>c01</td>
                        <td>c02</td>
                    </tr>
                    <tr>
                        <td>c10</td>
                        <td>c11</td>
                        <td>c12</td>
                    </tr>
                    <tr>
                        <td>c20</td>
                        <td>c21</td>
                        <td>c22</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
    <div id="archive">
        <div id="table2">
            <table>
                <tbody>
                    <tr>
                        <td>COL0</td>
                        <td>COL1</td>
                        <td>COL2</td>
                    </tr>
                    <tr>
                        <td>a00</td>
                        <td>a01</td>
                        <td>a02</td>
                    </tr>
                    <tr>
                        <td>a10</td>
                        <td>a11</td>
                        <td>a12</td>
                    </tr>
                    <tr>
                        <td>a20</td>
                        <td>a21</td>
                        <td>a22</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
    <div id="future">
        <div id="table3">
            <table>
                <tbody>
                    <tr>
                        <td>COL0</td>
                        <td>COL1</td>
                        <td>COL2</td>
                    </tr>
                    <tr>
                        <td>f00</td>
                        <td>f01</td>
                        <td>f02</td>
                    </tr>
                    <tr>
                        <td>f10</td>
                        <td>f11</td>
                        <td>f12</td>
                    </tr>
                    <tr>
                        <td>f20</td>
                        <td>f21</td>
                        <td>f22</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
