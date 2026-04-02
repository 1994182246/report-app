import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  AlignmentType, convertInchesToTwip, Table, TableRow, TableCell,
  WidthType, HeightRule, VerticalAlign,
} from 'docx';

const CHINESE_NUMS = ['一','二','三','四','五','六','七','八','九','十'];

function formatDateTitle(dateString: string) {
  const date = new Date(dateString);
  return `${date.getFullYear()}年${date.getMonth() + 1}月`;
}

function formatDateText(dateString: string) {
  const date = new Date(dateString);
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

function getDisplayItem(item: string) {
  if (item === '麻精药品' || item.includes('毒麻')) {
    return '麻醉、精神、未列管全身麻醉药品';
  }
  return item;
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const b64 = base64.split(',')[1] ?? base64;
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function getImageDimensions(dataURI: string): Promise<{ width: number; height: number }> {
  // In Node environment we can't use Image(), so parse from base64 header or default
  // We'll use a fixed max and let docx scale it
  return Promise.resolve({ width: 300, height: 300 });
}

export default async (req: Request) => {
  if (req.method !== 'POST') {
    return new Response('Method not allowed', { status: 405 });
  }

  let body: any;
  try {
    body = await req.json();
  } catch {
    return new Response('Invalid JSON', { status: 400 });
  }

  const { inspectionDate, inspectionDepartment, inspectionItem, problems, suggestions } = body;

  // Build image index map
  const imageIndexMap = new Map<string, string>();
  let imageCounter = 0;
  for (const problem of problems) {
    if (problem.image) {
      imageIndexMap.set(problem.id, CHINESE_NUMS[imageCounter] || (imageCounter + 1).toString());
      imageCounter++;
    }
  }

  const docChildren: any[] = [];

  // 标题
  docChildren.push(
    new Paragraph({
      children: [new TextRun({ text: `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告`, font: '宋体', size: 44, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );

  // 导语
  docChildren.push(
    new Paragraph({
      children: [new TextRun({
        text: `${formatDateText(inspectionDate)}，药学部药品质量管理工作小组对${inspectionDepartment}科室${getDisplayItem(inspectionItem)}进行检查，现将存在问题整理汇报如下：`,
        font: '宋体', size: 28,
      })],
      indent: { firstLine: 560 },
      spacing: { line: 560, lineRule: 'exact' as any, after: 200 },
    })
  );

  // 一、存在问题
  docChildren.push(
    new Paragraph({
      children: [new TextRun({ text: '一、存在问题：', font: '宋体', size: 28, bold: true })],
      spacing: { line: 560, lineRule: 'exact' as any, before: 200, after: 200 },
    })
  );

  for (let i = 0; i < problems.length; i++) {
    const problem = problems[i];
    const imgIndexStr = imageIndexMap.get(problem.id);
    const children: any[] = [
      new TextRun({ text: `${i + 1}. `, font: '宋体', size: 28 }),
      problem.department
        ? new TextRun({ text: problem.department, font: '宋体', size: 28, bold: true })
        : new TextRun({ text: '___', font: '宋体', size: 28 }),
      new TextRun({ text: `：${problem.description || '___'}`, font: '宋体', size: 28 }),
    ];
    if (problem.image && imgIndexStr) {
      children.push(new TextRun({ text: `（见图${imgIndexStr}）。`, font: '宋体', size: 28 }));
    }
    docChildren.push(
      new Paragraph({ children, indent: { firstLine: 560 }, spacing: { line: 560, lineRule: 'exact' as any } })
    );
  }

  // 图片表格
  const problemsWithImages = problems
    .filter((p: any) => p.image)
    .map((p: any) => ({ p, indexStr: imageIndexMap.get(p.id)! }));

  if (problemsWithImages.length > 0) {
    docChildren.push(new Paragraph({ text: '', spacing: { after: 200 } }));

    const colCount = problemsWithImages.length === 1 ? 1 : problemsWithImages.length === 2 ? 2 : 3;
    const colWidth = Math.floor(100 / colCount);
    const MAX_DIM = colCount === 1 ? 300 : colCount === 2 ? 240 : 180;

    const tableRows: any[] = [];
    for (let i = 0; i < problemsWithImages.length; i += colCount) {
      const chunk = problemsWithImages.slice(i, i + colCount);
      const imageCells: any[] = [];
      const captionCells: any[] = [];

      for (let j = 0; j < colCount; j++) {
        const item = chunk[j];
        if (item?.p?.image) {
          const imgBuffer = base64ToArrayBuffer(item.p.image);
          imageCells.push(new TableCell({
            width: { size: colWidth, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new ImageRun({ data: imgBuffer, transformation: { width: MAX_DIM, height: MAX_DIM }, type: 'png' })],
            })],
          }));
          captionCells.push(new TableCell({
            width: { size: colWidth, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `图${item.indexStr}`, font: '宋体', size: 24, bold: true })],
            })],
          }));
        } else {
          imageCells.push(new TableCell({ width: { size: colWidth, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: '' })] }));
          captionCells.push(new TableCell({ width: { size: colWidth, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: '' })] }));
        }
      }

      tableRows.push(new TableRow({ height: { value: 2952, rule: HeightRule.EXACT }, children: imageCells }));
      tableRows.push(new TableRow({ height: { value: 400, rule: HeightRule.ATLEAST }, children: captionCells }));
    }

    docChildren.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }));
  }

  // 二、改进建议
  docChildren.push(
    new Paragraph({
      children: [new TextRun({ text: '二、改进建议：', font: '宋体', size: 28, bold: true })],
      spacing: { line: 560, lineRule: 'exact' as any, before: 200, after: 200 },
    })
  );

  for (let i = 0; i < suggestions.length; i++) {
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: `${i + 1}. ${suggestions[i].text || '___'}`, font: '宋体', size: 28 })],
        indent: { firstLine: 560 },
        spacing: { line: 560, lineRule: 'exact' as any },
      })
    );
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: convertInchesToTwip(1.45),
            bottom: convertInchesToTwip(1.37),
            left: convertInchesToTwip(1.1),
            right: convertInchesToTwip(1.02),
          },
        },
      },
      children: docChildren,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  const fileName = `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告.docx`;

  return new Response(buffer, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
      'Access-Control-Allow-Origin': '*',
    },
  });
};

export const config = { path: '/api/generate-docx' };
