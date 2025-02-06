Office.onReady();

/** inserts a signature automatically when a new message is composed or a recipient is changed */
function insertSignature(event) {
  Office.context.mailbox.item.body.setSignatureAsync(`
  <table style="width:545px; border-collapse:collapse">
    <tbody>
      <tr>
        <td style="font-size:9pt; font-family:Arial; width:400px; color:#3c3c3b; padding:0 0 5px 5px; border-left:#0099cc 2px solid" valign="top">
          <b><span style="font-size:12pt; color:#3c3c3b">${Math.random().toString(36).slice(2, 7)} ${Math.random().toString(36).slice(2, 7)}</span></b><br>
          <em style="font-size:9pt; color:#3c3c3b">Manager</em>
        </td>
      </tr>
      <tr>
        <td style="font-size:9pt; font-family:Arial; width:400px; padding:0 0 5px 5px; border-left:#0099cc 2px solid" valign="top">
          <strong style="font-size:12pt; color:#0099cc">Your Company</strong><br>
          <span style="color:#0099cc">p:</span> <a style="text-decoration:none; color:#3c3c3b" href="tel:425-555-0100">425-555-0100</a><br>
          <span style="color:#0099cc">a:</span> <font color="#3d3c3f">22 Branding Blvd, Azure Hill, NV, 89404, USA</font><br>
          <span style="color:#0099cc">w:</span> <a style="text-decoration:none; color:#3c3c3b" href="http://www.yourdomain.url/">www.yourdomain.url</a><br>
          <span style="color:#0099cc">e:</span> <a style="text-decoration:none; color:#0099cc" href="mailto:admin@M365x84368197.OnMicrosoft.com">admin@M365x84368197.OnMicrosoft.com</a>
        </td>
      </tr>
    </tbody>
  </table>        
`, { coercionType: "html" }, function () { event.completed(); });
}

Office.actions?.associate("insertSignature", insertSignature);
