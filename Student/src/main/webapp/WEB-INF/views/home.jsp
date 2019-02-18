<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
	pageEncoding="ISO-8859-1"%>

<%@taglib prefix="c" uri="http://java.sun.com/jstl/core_rt"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="ISO-8859-1">
<title>Insert title here</title>

<link rel="stylesheet"
	href="${pageContext.request.contextPath}/css/bootstrap.min.css">
<script type="text/javascript"
	src="${pageContext.request.contextPath}/js/jquery.min.js"></script>
<script type="text/javascript"
	src="${pageContext.request.contextPath}/js/bootstrap.min.js"></script>
	
	<c:set var="a" value="" />
	
	
	<script type="text/javascript">
	$(document).ready(function() {
		$('#sbtn').bind('click',save);
	})
	
	function save() {
		
		var data=$('#formid').serialize();
		
		$.ajax({
			method:'POST',
			url:'saveStudent',
			data:data,
			
			success:function(){
				alert('saved successfully...');
				url="";
				window.location.href="${pageContext.request.contextPath}/";
			},
			
			error:function(){
				alert('error while saving...');
			}
			
		});
		
	}
	
	function edit(val) {
		
		$.ajax({
			method:'GET',
			url:'editStudent?id='+val,
			
			success:function(response){
				$('#id').val(response.id);
				$('#name').val(response.name);
				$('#place').val(response.place);
				$('#email').val(response.email);
				//'${a}'=response.id;
			},
			
			error:function(){
				alert('error while retriving...');
			}
			
		});
		
	}
	
	function deleteStudent(val) {
		window.location.href="deleteStudent?id="+val;
	}
	
	</script>

</head>
<body>

	<div class="container">
		<div class="row">

			<div class="col-md-offset-2 col-md-8">

				<div class="panel panel-primary">

					<div class="panel-heading">
						<h4>Student Registration</h4>
					</div>

					<div class="panel-body">

						<form id="formid"
							class="form-horizontal">

							<div class="form-group">

								<label class="control-label col-md-4" for="name">Student
									Name:</label>
								<div class="col-md-8">
									<input type="text" name="name" id="name"
										placeholder="Enter Your Name" class="form-control" />
								</div>

							</div>

							<div class="form-group">

								<label class="control-label col-md-4" for="place">Student
									Place:</label>
								<div class="col-md-8">
									<input type="text" name="place" id="place"
										placeholder="Enter Your Place" class="form-control" />
								</div>

							</div>

							<div class="form-group">

								<label class="control-label col-md-4" for="email">Student
									Email:</label>
								<div class="col-md-8">
									<input type="text" name="email" id="email"
										placeholder="Enter Your Email" class="form-control" />
								</div>

							</div>
							
							<%-- <c:if test=${not empty a}> --%>
							<input type="hidden" name="id" id="id" value="0" />
							<%-- </c:if> --%>
							
							<div class="form-group">

								<div class="col-md-8">
									<input type="button" value="Register" id="sbtn" />
								</div>

							</div>

						</form>

							<div class="form-group">

								<div class="col-md-8">
									<a href="excelrep.html"><input type="button" value="Report"  /></a>
								</div>

							</div>
					</div>
				</div>
			</div>

		</div>


		<h2>All Student Record</h2>

		<table class="table table-bordered table-striped">
			<thead>
				<tr>
					<th>SL.NO</th>
					<th>NAME</th>
					<th>PLACE</th>
					<th>EMAIL</th>
					<th>EDIT</th>
					<th>DELETE</th>
				</tr>
			</thead>

			<tbody>

				<c:forEach items="${students}" var="s" varStatus="loop">
					<tr>
						<td>${loop.count }</td>
						<td>${s.name }</td>
						<td>${s.place }</td>
						<td>${s.email }</td>
						<td>
							<button onclick="edit(${s.id});">EDIT</button>
						</td>
						<td>
							<button onclick="deleteStudent(${s.id});">DELETE</button>
						</td>

					</tr>
				</c:forEach>




			</tbody>

		</table>



	</div>

</body>
</html>