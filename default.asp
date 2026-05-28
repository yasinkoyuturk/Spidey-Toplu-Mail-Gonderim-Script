<!-- #include file="ust.asp" -->
<%
'// Kodlama: Yasin Koyutürk
'// Konu   : Toplu mail gönderimi
'Eðer iþlem boþ geliyorsa formu gösteriyoruz
if request("islem")="" then
%>
	<form method="POST" action="?islem=gonder">
	<table border="0" cellpadding="0" cellspacing="0" width="400" id="table1">
		<tr>
			<td width="121"><span class="normal_ryazi">Gönderen mail</span></td>
			<td width="8"><span class="normal_ryazi">:</span></td>
			<td> 
			<input type="text" name="gonderen" size="28" value="" class="normal_yazi"></td>
		</tr>
		<tr>
			<td width="121"><span class="normal_ryazi">Gönderen isim</span></td>
			<td width="8"><span class="normal_ryazi">:</span></td>
			<td> 
			<input type="text" name="gonderen_ad" size="28" value="" class="normal_yazi"></td>
		</tr>
		<tr>
			<td width="121"><span class="normal_ryazi">Konu</span></td>
			<td width="8"><span class="normal_ryazi">:</span></td>
			<td> 
			<input type="text" name="konu" size="28" value="" class="normal_yazi"></td>
		</tr>
		<tr>
			<td width="121"><span class="normal_ryazi">Mesajýnýz:<br>
			(HTML desteði mevcut)</span></td>
			<td width="8"><span class="normal_ryazi">:</span></td>
			<td> <textarea rows="15" name="msj" cols="49" class="normal_yazi"></textarea></td>
		</tr>
		<tr>
			<td width="121">&nbsp;</td>
			<td width="8">&nbsp;</td>
			<td>
			<input type="submit" value="Gönder" name="B1" class="normal_yazi"></td>
		</tr>
	</table>
	</form>
<%
'Eðer iþlem gönder ise
elseif request("islem")="gonder" then

	'Veritabanýnýzýn;
	' id(otomatik sayý)
	' ad(metin)
	' mail(not)
	' þeklinde üc sütun bulundurduðunu ve isminin vt.mdb olduðunu kabul ediyoruz.
	Set bag = Server.Createobject("ADODB.Connection")
	bag.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("vt.mdb"))
	
	'Recordsetimizi yazýyoruz
	' uyeler tablosundaki reklam_maillerini kabul etmiþ üyeler için seçiyoruz
	Set uyelik = Server.Createobject("ADODB.Recordset")
	uyelik.Open "SELECT * FROM uyeler where reklam_mailleri=True" ,bag,1,3

	'Reklam maillerini kabul etmiþ üyelerin hepsine yollayacak döngümüzü baþlatýyoruz
	for i=1 to uyelik.recordcount
	
	'Eðer uyelik recordsetimiz boþ ise döngüden çýkýyoruz
	if uyelik.eof then exit for

		'Persist bileþeni ile yollayacaðýz
		Set Mail = Server.CreateObject("Persits.MailSender")

		'Sitenizin mail sunucusu ör: mail.oyuncum.net
		' Hostunuz localhost ismi ile mail yollamanýza izin veriyorsa formda
		' verdiðiniz herhangi bir isim ve mail ile yollayabilirsiniz
		' örneðin arkadaþýnýzýn haberi olmadan onun mail adresi ve ismiyle bile =)
		Mail.Host = "localhost"
	
		'Kullanýcý adý ve þifre kýsmý
		' Eðer siteniz desteklemiyor ve host kýsmýný localhost deðilde 
		' mail sunucuzu girdiyseniz alttaki iki satýrýn baþýndaki
		' týrnaklarý kaldýrýn.
		'Mail.Username = "agi82@agi82.net" 'posta hesabýnýzýn kul adi
	    'Mail.Password = "*********" 'posta hesabýnýzýn Þifresi
    		
		'Formdan gelen, maili yollayanýn adresi
		Mail.From = Request.form("gonderen")

		'Formdan gelen Gönderici adý
		Mail.FromName = Request.form("gonderen_ad")
		
		'Veritabanýndan gelen üye mail adresi
		Mail.AddAddress uyelik("mail")
		
		'Mailin konusu
		Mail.Subject = Request.form("konu")

		'HTML destekleyip desteklemediðini giriyoruz.Kapatmak için False yapýn
		Mail.isHTML = True
		
		'Formdan gelen msj kýsmý
		Mail.Body = "Kime:"&uyelik("ad")&"<br>Msj:"&Request.form("msj")
		
		'Bir hata olursa hata açýklamasýný görüntüle
		On Error Resume Next
		
		'Gönderiyoruz
		Mail.Send 
		
	
	'Döngü sonunda bir sonraki kayda geçiyoruz ve devam ediyoruz
	uyelik.movenext
	
	'Döngüyü bitiriyoruz
	next
		'Eðer hata varsa açýklamasýný yazýyoruz, yoksa gönderildi yazdýrýyoruz
		If Err <> 0 Then 
			strErr = Err.Description
			Response.Write "<FONT COLOR=""#FF0000"">Hata: <I>"& strErr &"</I></FONT>"
		else
			bSuccess = True
			Response.Write "<FONT COLOR=""#00A000"">Mesajýnýz  "&i&" kiþiye gönderildi.</FONT>"
		End If

end if
%><!-- #include file="alt.asp" -->