function main(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet("Sheet1");
  const range = sheet.getRange("A2:D9");
  const currentDate = new Date();
  const values: (number | string | boolean)[][] = range.getValues();
  const dueThisWeek: Array<string> = [
    `    <table><thead><tr><th>Date</th><th>Module Code</th><th>Assassment</th>
    <th>Venue</th></tr></thead><tbody>`,
  ];

  for (let i = 0; i < values.length; i++) {
    const jsDate = convertDate(values[i][0]);
    const milliseconds: number = jsDate - currentDate;
    const days = milliseconds / (24 * 60 * 60 * 1000);
    if (days > 3 || days < 1) {
      continue;
    }

    const task: string = `
      <tr><td>${jsDate}</td><td>${values[i][1]}</td>
      <td>${values[i][2]}</td><td>${values[i][3]}</td></tr>
      `;

    dueThisWeek.push(task);
  }

  dueThisWeek.push(`
      </tbody>
    </table>
    `);

  return dueThisWeek.toString().replace(",", " ");
}

function convertDate(excelDateValue: number) {
  return new Date(Math.round((excelDateValue - 25_569) * 86_400 * 1000));
}
