import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  AlignmentType, convertInchesToTwip, Table, TableRow, TableCell,
  WidthType, HeightRule, VerticalAlign,
} from 'docx';
import { getStore } from '@netlify/blobs';

const CHINESE_NUMS = ['一','二','三','四','五','六','七','八','九','十'];

function formatDateTitle(d: string) {
  const date = new Date(d);
  return `${date.getFullYear()}年${date.getMonth() + 1}月`;
}
function formatDateText(d: string) {
  const date = new Date(d);
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}
function getDisplayItem(item: string) {
  if (item === '麻精药品' || item.includes('毒麻')) return '麻醉、精神、未列管全身麻醉药品';
  return item;
}
function base64ToBuffer(base64: string): ArrayBuffer {
  const b64 = base64.includes(',') ? base64.split(',')[1] : base64;
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

export default async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST', 'Access-Control-Allow-Headers': 'Content-Type' } });
  }
  if (req.method !== 'POST') return new Response('Method not allowed', { status: 405 });

  let body: any;
  try { body = await req.json(); } catch { return new Response('Invalid JSON', { status: 400 }); }

  const { inspectionDate, inspectionDepartment, inspectionItem, problems, suggestions } = body;

  const imageIndexMap = new Map<string, string>();
  let imgCounter = 0;
  for (const p of problems) {
    if (p.image) imageIndexMap.set(p.id, CHINESE_NUMS[imgCounter++] || String(imgCounter));
  }

  const children: any[] = [];

  children.push(new Paragraph({
    children: [new TextRun({ text: `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告`, font: '宋体', size: 44, bold: true })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
  }));

  children.push(new Paragraph({
    children: [new TextRun({ text: `${formatDateText(inspectionDate)}，药学部药品质量管理工作小组对${inspectionDepartment}科室${getDisplayItem(inspectionItem)}进行检查，现将存在问题整理汇报如下：`, font: '宋体', size: 28 })],
    indent: { firstLine: 560 },
    spacing: { line: 560, lineRule: 'exact' as any, after: 200 },
  }));

  children.push(new Paragraph({
    children: [new TextRun({ text: '一、存在问题：', font: '宋体', size: 28, bold: true })],
    spacing: { line: 560, lineRule: 'exact' as any, before: 200, after: 200 },
  }));

  for (let i = 0; i < problems.length; i++) {
    const p = problems[i];
    const imgIdx = imageIndexMap.get(p.id);
    const runs: any[] = [
      new TextRun({ text: `${i + 1}. `, font: '宋体', size: 28 }),
      p.department ? new TextRun({ text: p.department, font: '宋体', size: 28, bold: true }) : new TextRun({ text: '___', font: '宋体', size: 28 }),
      new TextRun({ text: `：${p.description || '___'}`, font: '宋体', size: 28 }),
    ];
    if (p.image && imgIdx) runs.push(new TextRun({ text: `（见图${imgIdx}）。`, font: '宋体', size: 28 }));
    children.push(new Paragraph({ children: runs, indent: { firstLine: 560 }, spacing: { line: 560, lineRule: 'exact' as any } }));
  }

  const withImages = problems.filter((p: any) => p.image).map((p: any) => ({ p, indexStr: imageIndexMap.get(p.id)! }));
  if (withImages.length > 0) {
    children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
    const cols = withImages.length === 1 ? 1 : withImages.length === 2 ? 2 : 3;
    const colW = Math.floor(100 / cols);
    const maxDim = cols === 1 ? 300 : cols === 2 ? 240 : 180;
    const rows: any[] = [];
    for (let i = 0; i < withImages.length; i += cols) {
      const chunk = withImages.slice(i, i + cols);
      const imgCells: any[] = [], capCells: any[] = [];
      for (let j = 0; j < cols; j++) {
        const item = chunk[j];
        if (item?.p?.image) {
          imgCells.push(new TableCell({ width: { size: colW, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: { top: 100, bottom: 100, left: 100, right: 100 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: base64ToBuffer(item.p.image), transformation: { width: maxDim, height: maxDim }, type: 'png' })] })] }));
          capCells.push(new TableCell({ width: { size: colW, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `图${item.indexStr}`, font: '宋体', size: 24, bold: true })] })] }));
        } else {
          imgCells.push(new TableCell({ width: { size: colW, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: '' })] }));
          capCells.push(new TableCell({ width: { size: colW, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: '' })] }));
        }
      }
      rows.push(new TableRow({ height: { value: 2952, rule: HeightRule.EXACT }, children: imgCells }));
      rows.push(new TableRow({ height: { value: 400, rule: HeightRule.ATLEAST }, children: capCells }));
    }
    children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows }));
  }

  children.push(new Paragraph({
    children: [new TextRun({ text: '二、改进建议：', font: '宋体', size: 28, bold: true })],
    spacing: { line: 560, lineRule: 'exact' as any, before: 200, after: 200 },
  }));
  for (let i = 0; i < suggestions.length; i++) {
    children.push(new Paragraph({
      children: [new TextRun({ text: `${i + 1}. ${suggestions[i].text || '___'}`, font: '宋体', size: 28 })],
      indent: { firstLine: 560 },
      spacing: { line: 560, lineRule: 'exact' as any },
    }));
  }

  const doc = new Document({
    sections: [{ properties: { page: { margin: { top: convertInchesToTwip(1.45), bottom: convertInchesToTwip(1.37), left: convertInchesToTwip(1.1), right: convertInchesToTwip(1.02) } } }, children }],
  });

  const buffer = await Packer.toBuffer(doc);
  const uint8 = new Uint8Array(buffer).buffer; // ensure plain ArrayBuffer for Blobs API
  const fileName = `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告.docx`;

  const store = getStore({ name: 'docx-temp', consistency: 'strong' });
  const key = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  await store.set(key, uint8, { metadata: { fileName } });

  // 返回下载 URL（通过 /api/download 路由提供）
  const downloadUrl = new URL(req.url);
  const fileUrl = `${downloadUrl.origin}/api/download-docx?key=${key}&name=${encodeURIComponent(fileName)}`;

  return new Response(JSON.stringify({ url: fileUrl, fileName }), {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
  });
};

export const config = { path: '/api/generate-docx' };
