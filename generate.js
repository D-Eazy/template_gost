function generate() {
    let t1 = document.getElementById('t1').value;
    let t2 = document.getElementById('t2').value;
    let t3 = document.getElementById('t3').value;
    let t4 = document.getElementById('t4').value;
    let t5 = document.getElementById('t5').value;
    let t6 = document.getElementById('t6').value;
    let t7 = document.getElementById('t7').value;

    let res1, res2;

    let item1 = document.getElementsByTagName('input')[0];
    if (item1.getAttribute('type') === "checkbox"
        && item1.checked
        && item1.name === "s1") {
            res1 = "Министерство образования и науки РФ"
    } else {
        res1 = '';
    }

    let item2 = document.getElementsByTagName('input')[1];
    if (item1.getAttribute('type') === "checkbox"
        && item2.checked
        && item2.name === "s2") {
        res2 = "Федеральное государственное бюджетное общеобразовательное учреждение высшего образования «Ивановский государственный энергетический университет имени В.И. Ленина»"
    } else {
        res2 = '';
    }

    if ((t1 === '') || (t2 === '') || (t3 === '') || (t4 === '') || (t5 === '') || (t6 === '') || (t7 === '')) {
        alert("Заполните все поля!");
    } else {
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    new docx.Paragraph({
                        frame: {
                            position: {
                                x: 1000,
                                y: 3000,
                            },
                            width: 4000,
                            height: 1000,
                            anchor: {
                                horizontal: FrameAnchorType.MARGIN,
                                vertical: FrameAnchorType.MARGIN,
                            },
                            alignment: {
                                x: HorizontalPositionAlign.CENTER,
                                y: VerticalPositionAlign.TOP,
                            },
                        },
                        children: [
                            new docx.TextRun({
                                text: res1,
                                bold: true,
                                size: 26,
                                alignment: AlignmentType.CENTER,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: res2,
                                bold: true,
                                size: 26,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Кафедра " + t1,
                                bold: true,
                                size: 26,
                                break: 5,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Дисциплина: " + t2,
                                bold: true,
                                size: 26,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Отчет по лабораторной работе на тему: " + t3,
                                bold: true,
                                size: 26,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Работу выполнил:",
                                bold: true,
                                size: 24,
                                break: 5,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Студент группы " + t4,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: t5,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Проверил:",
                                bold: true,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: t6,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Иваново " + t7,
                                size: 24,
                                break: 5,
                            }),
                        ],
                    }),
                ],
            }]
        });

        docx.Packer.toBlob(doc).then(blob => {
            console.log(blob);
            saveAs(blob, "template.docx");
            console.log("Документ создан");
        });
    }
}