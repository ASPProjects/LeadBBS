	<script type="text/javascript">
	<!--
	function surfto1(list)
	{
		var myindex1  = list.selectedIndex;
		if (myindex1 != 0)
		{
			var URL = "../" + list.options[list.selectedIndex].value;
			this.location.href = URL; 
			target = '_self';
		}
	}
	-->
	</script>
	<select name="jumpto" onchange="surfto1(this)" style="width:100px;">
		<option value="Boards.asp">�л����桭</option>
		<option value="Boards.asp">��̳��ҳ</option>
		<option value="Boards.asp?Assort=100">��Default</option>
		<option value="b/b.asp?B=100">��Default Forum</option>
		<option value="b/b.asp?B=444">��Recycle</option>
	</select>
