<?php
/** Database Connection file * */
include 'conn.php'
?>

<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="robots" content="noindex">

    <title>TVC Report Generator</title>

    <!-- Bootstrap core CSS -->
    <link rel="canonical" href="http://bootstrapformhelpers.com/timepicker/"/>
    <link href="dist/css/bootstrap.css" rel="stylesheet">
    <link href="dist/css/bootstrap-formhelpers.min.css" rel="stylesheet">
    <link href="dist/css/datepicker.css" rel="stylesheet">
    <!-- Custom styles for this template -->
    <link href="dist/css/starter-template.css" rel="stylesheet">
    <link href="dist/css/prettify.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" media="screen" href="dist/css/bootstrap-datetimepicker.min.css">
    <!-- Just for debugging purposes. Don't actually copy this line! -->
    <!--[if lt IE 9]><script src="../../docs-assets/js/ie8-responsive-file-warning.js"></script><![endif]-->

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
  </head>

  <body>

    <div class="navbar navbar-inverse navbar-fixed-top" role="navigation">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
            <span class="sr-only">Toggle navigation</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <a class="navbar-brand" href="#">TVC iPad Application Report Generator</a>
        </div>
        <div class="collapse navbar-collapse">
          <ul class="nav navbar-nav">
            <!--<li class="active"><a href="#">Home</a></li>-->
            <!--<li><a href="#about">About</a></li>
            <li><a href="#contact">Contact</a></li>-->
          </ul>
        </div><!--/.nav-collapse -->
      </div>
    </div>

    <div class="container">

      <div class="starter-template">
        <!--<h1>Report Generator</h1>-->
        <p class="lead">Select the dates to generate report</p
        <?php 
        if(isset($_GET['error']))
        {
            echo '<div class="well"><div class="alert alert-danger" id="alert">
				<strong>No Data found</strong>
			  </div></div>';
        }
            ?>
        <div class="well">
            
            
            <div class="alert alert-danger" id="alert" style="display:none">
				<strong></strong>
			  </div>
            <form action="export.php" method="post" onsubmit="return checkDate()">
                <select name="opt" id="opt" onchange="disp()">
                    <option value="boneyard">Boneyard Report</option>
                    <option value="pandl">P and L Report</option>
                    <option value="spandl">Stand P and L Report</option>
                    <option value="orderrep">Order Report</option>
                    <option value="prodrep">Product Report</option>
                    <option value="voidrep">VOID Report</option>
                    <option value="locordrep">Location Order Report</option>
                    <option value="prodlist">Product List</option>
                    <option value="loclist">Location List</option>
                    <option value="boneyardcountdown">Boneyard Count Down</option>
                </select>
			<table class="table" id="dispDate">
				<thead>
					<tr>
                                            <th style="text-align: center">Start date &nbsp;<a href="#" class="btn btn-default" id="dp4" data-date-format="yyyy-mm-dd" data-date="<?php echo date("Y-m-d") ?>">Change</a></th>
                                            <th style="text-align: center">End date &nbsp;<a href="#" class="btn btn-default" id="dp5" data-date-format="yyyy-mm-dd" data-date="<?php echo date("Y-m-d") ?>">Change</a></th>
					</tr>
				</thead>
				<tbody>
					<tr>
                                            <td id="startDate" style="font-weight: bold; text-align: center" ><?php echo date("Y-m-d") ?></td>
                                            <td id="endDate" style="font-weight: bold; text-align: center"><?php echo date("Y-m-d") ?></td>
					</tr>
                                        <tr>                                                                                
                                           
					</tr>
                                        <tr style="width: ">
                                            <td colspan="2">
                                        </td>
                                        </tr>
				</tbody>
			</table>
                <div id="hide" style="display: none">
                <div id="st1" style="float: left">Start Time:</div> <div id="t1" class="bfh-timepicker" style="width: 300px; padding-left: 138px"></div><br>
                <div id="st2" style="float: left">End Time:</div> <div id="t2" class="bfh-timepicker" style="width: 300px; padding-left: 138px"></div><br>
                 
                <div id="locHide" style="display: none"><div  style="float: left;">Select Location: </div> <div style="width: 145px; float: left">
                 <select name="location">
                    <?php 
                    $stmt = $conn->prepare('SELECT NAME FROM `location` WHERE CHARGEABLE = :ch AND NAME != :loc');
                    $stmt->execute(array(':ch'=>'Chargeable',':loc'=>'Boneyard'));
                    $order_id_arr = array();
                    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
                        echo "<option>{$row['NAME']}</option>";
                    }
                    ?>
                </select>
                     </div>
                </div>
                
                
                
                 <div style="clear: both" ></div>
                </div>
                <div id="loc1Hide" style="display: none"><div  style="float: left;">Select Location: </div> <div style="width: 145px; float: left">
                 <select name="location1">
                    <?php 
                    $stmt = $conn->prepare('SELECT NAME FROM `location` WHERE NAME != :loc');
                    $stmt->execute(array(':loc'=>'Boneyard'));
                    $order_id_arr = array();
                    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
                        echo "<option>{$row['NAME']}</option>";
                    }
                    ?>
                </select>
                     </div>
                </div>
                 <br>                 
                <input type="hidden" name="dateFrom" id="dateFrom" value="">
                <input type="hidden" name="dateTo" id="dateTo" value="">
                <input type="hidden" name="timeFrom" id="timeFrom" value="">
                <input type="hidden" name="timeTo" id="timeTo" value="">
                <select name="fileType"><option value="excel">Excel</option><option value="pdf">PDF</option></select>
                <input type="submit" id="but" style="text-align: center" class="btn btn-default">
            </form>
          </div>
      </div>

    </div><!-- /.container -->


    <!-- Bootstrap core JavaScript
    ================================================== -->
    <!-- Placed at the end of the document so the pages load faster -->
    <script src="dist/js/jquery-1.10.2.min.js"></script>
    <script src="dist/js/bootstrap.min.js"></script>
    <script src="dist/js/bootstrap.min1.js"></script>
    <script src="dist/js/bootstrap-datepicker.js"></script>
    <script src="dist/js/bootstrap-formhelpers.min.js"></script>
    
    <script type="text/javascript" src="dist/js/bootstrap-datetimepicker.min.js"></script>
    <script type="text/javascript" src="dist/js/bootstrap-datetimepicker.pt-BR.js"></script> 
    
    <script src="dist/js/prettify.js"></script>
	<script>
	if (top.location != location) {
    top.location.href = document.location.href ;
  }
		$(function(){
			window.prettyPrint && prettyPrint();
			$('#dp1').datepicker({
				format: 'mm-dd-yyyy'
			});
			$('#dp2').datepicker();
			$('#dp3').datepicker();
			$('#dp3').datepicker();
			$('#dpYears').datepicker();
			$('#dpMonths').datepicker();
			
			
			var startDate = new Date();
			var endDate = new Date();
			$('#dp4').datepicker()
				.on('changeDate', function(ev){
					if (ev.date.valueOf() > endDate.valueOf()){
						$('#alert').show().find('strong').text('The start date can not be greater then the end date');
					} else {
						$('#alert').hide();
						startDate = new Date(ev.date);
						$('#startDate').text($('#dp4').data('date'));
					}
					$('#dp4').datepicker('hide');
				});
			$('#dp5').datepicker()
				.on('changeDate', function(ev){
					if (ev.date.valueOf() < startDate.valueOf()){
						$('#alert').show().find('strong').text('The end date can not be less then the start date');
					} else {
						$('#alert').hide();
						endDate = new Date(ev.date);
						$('#endDate').text($('#dp5').data('date'));
					}
					$('#dp5').datepicker('hide');
				});

        // disabling dates
        var nowTemp = new Date();
        var now = new Date(nowTemp.getFullYear(), nowTemp.getMonth(), nowTemp.getDate(), 0, 0, 0, 0);

        var checkin = $('#dpd1').datepicker({
          onRender: function(date) {
            return date.valueOf() < now.valueOf() ? 'disabled' : '';
          }
        }).on('changeDate', function(ev) {
          if (ev.date.valueOf() > checkout.date.valueOf()) {
            var newDate = new Date(ev.date)
            newDate.setDate(newDate.getDate() + 1);
            checkout.setValue(newDate);
          }
          checkin.hide();
          $('#dpd2')[0].focus();
        }).data('datepicker');
        var checkout = $('#dpd2').datepicker({
          onRender: function(date) {
            return date.valueOf() <= checkin.date.valueOf() ? 'disabled' : '';
          }
        }).on('changeDate', function(ev) {
          checkout.hide();
        }).data('datepicker');
		});
                
            function checkDate()
            {
                var startDate = document.getElementById('startDate').innerText;
                var endDate = document.getElementById('endDate').innerText;
                document.getElementById('dateFrom').value = startDate;
                document.getElementById('dateTo').value = endDate;
                var time = document.getElementsByClassName('form-control');
 
                document.getElementById('timeFrom').value = time[0].value+":00";
                document.getElementById('timeTo').value = time[3].value+":00";
                //return false;
            }
            function disp()
            {
                var val = document.getElementById("opt").value;
                if(val == "pandl" || val == "spandl" ) {
                    document.getElementById('hide').style.display = "block";
                    document.getElementById('t1').style.display = "none";
                    document.getElementById('t2').style.display = "none";
                    document.getElementById('st1').style.display = "none";
                    document.getElementById('st2').style.display = "none";
                    if(val == "spandl") {
                        document.getElementById('locHide').style.display = "block";
                    } else {
                        document.getElementById('locHide').style.display = "none";
                        document.getElementById('loc1Hide').style.display = "none";
                    }
                } else if(val == "locordrep") {
                        document.getElementById('hide').style.display = "none";
                        document.getElementById('locHide').style.display = "none";
                        document.getElementById('loc1Hide').style.display = "block";
                } else {
                        document.getElementById('hide').style.display = "none";
                        document.getElementById('locHide').style.display = "none";
                        document.getElementById('loc1Hide').style.display = "none";
                }

                if(val == "loclist" || val == "prodlist" || val == "boneyardcountdown") {
                    document.getElementById('dispDate').style.display = "none";
                    document.getElementById('but').style.marginTop = "10px";
                } else {
                    document.getElementById('dispDate').style.display = "";
                    document.getElementById('but').style.marginTop = "0px";
                }
            }
            $('#datetimepicker3').datetimepicker({
                                                  pickDate: false
                                                });
           //$('#datetimepicker3').datetimepicker('show');
           
           $('#datetimepicker4').datetimepicker({
                                                  pickDate: false
                                                });
           //$('#datetimepicker4').datetimepicker('show');
	</script>
  </body>
</html>