import React, { useState, useRef } from 'react';
import { Plus, Trash2, Upload, FileText, Download } from 'lucide-react';
import { format } from 'date-fns';
import { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, convertInchesToTwip } from 'docx';
import { saveAs } from 'file-saver';

type Problem = {
  id: string;
  department: string;
  description: string;
  image: string | null;
};

type Suggestion = {
  id: string;
  text: string;
};

const BUILT_IN_SUGGESTIONS = [
  '加强人员培训，提高安全用药意识',
  '严格落实查对制度，确保账物相符',
  '完善药品管理登记本，记录详实',
  '增加专项检查频次，督促持续整改',
];

const App = () => {
  const [inspectionDate, setInspectionDate] = useState(format(new Date(), 'yyyy-MM-dd'));
  const [inspectionDepartment, setInspectionDepartment] = useState('内科');
  const [inspectionItem, setInspectionItem] = useState('高警示药品和抢救车药品');
  
  const [problems, setProblems] = useState<Problem[]>([
    { id: crypto.randomUUID(), department: '', description: '', image: null }
  ]);
  
  const [suggestions, setSuggestions] = useState<Suggestion[]>([
    { id: crypto.randomUUID(), text: '' }
  ]);

  const [suggestionTemplates, setSuggestionTemplates] = useState<string[]>(BUILT_IN_SUGGESTIONS);
  const [newTemplateText, setNewTemplateText] = useState('');
  const [selectedTemplate, setSelectedTemplate] = useState<string>('');

  const reportRef = useRef<HTMLDivElement>(null);

  const handleAddProblem = () => {
    setProblems([...problems, { id: crypto.randomUUID(), department: '', description: '', image: null }]);
  };

  const handleRemoveProblem = (id: string) => {
    setProblems(problems.filter(p => p.id !== id));
  };

  const handleProblemChange = (id: string, field: keyof Problem, value: string | null) => {
    setProblems(problems.map(p => p.id === id ? { ...p, [field]: value } : p));
  };

  const handleImageUpload = (id: string, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        handleProblemChange(id, 'image', reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleAddSuggestion = (text = '') => {
    setSuggestions([...suggestions, { id: crypto.randomUUID(), text }]);
  };

  const handleRemoveSuggestion = (id: string) => {
    setSuggestions(suggestions.filter(s => s.id !== id));
  };

  const handleSuggestionChange = (id: string, text: string) => {
    setSuggestions(suggestions.map(s => s.id === id ? { ...s, text } : s));
  };

  const handleAddTemplate = () => {
    if (newTemplateText.trim() && !suggestionTemplates.includes(newTemplateText.trim())) {
      setSuggestionTemplates([...suggestionTemplates, newTemplateText.trim()]);
      setNewTemplateText('');
    }
  };

  const handleRemoveTemplate = () => {
    if (selectedTemplate) {
      setSuggestionTemplates(suggestionTemplates.filter(t => t !== selectedTemplate));
      setSelectedTemplate('');
    }
  };

  const exportAsWord = async () => {
    // 辅助函数：将 Base64 数据 URI 转换为 ArrayBuffer
    const base64DataURLToArrayBuffer = (dataURI: string) => {
      const base64String = dataURI.split(',')[1];
      const binaryString = window.atob(base64String);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      return bytes.buffer;
    };

    // 辅助函数：获取图片的实际宽高比，用于在 Word 中计算合适尺寸
    const getImageDimensions = (dataURI: string): Promise<{ width: number, height: number }> => {
      return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
          resolve({ width: img.width, height: img.height });
        };
        img.src = dataURI;
      });
    };

    const docChildren: any[] = [];

    // 1. 标题 (方正小标宋简体，二号)
    // 二号字体大小为 22pt，在 docx 中大小单位是半点(half-points)，所以是 44
    docChildren.push(
      new Paragraph({
        text: `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告`,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: {
          after: 400,
        },
        style: "TitleStyle",
      })
    );

    // 2. 导语正文 (仿宋_GB2312，三号，首行缩进两字符)
    // 三号字体为 16pt -> size: 32
    // 首行缩进两字符：三号字宽 16pt，两字符即 32pt。换算为 twip (1 pt = 20 twip)，即 640 twips
    // 行距：单倍行距或固定值，这里设置常用的 1.5 倍行距或特定固定值 (28磅 -> 560)
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${formatDateText(inspectionDate)}，药学部药品质量管理工作小组对${inspectionDepartment}科室${getDisplayItem(inspectionItem)}进行检查，现将存在问题整理汇报如下：`,
            font: "仿宋_GB2312",
            size: 32, // 三号
          }),
        ],
        indent: {
          firstLine: 640,
        },
        spacing: {
          line: 560, // 行距 28 磅
          lineRule: "exact",
          after: 200,
        },
      })
    );

    // 3. 一、存在问题
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "一、存在问题：",
            font: "黑体", // 公文一级标题常使用黑体
            size: 32,
            bold: true,
          }),
        ],
        spacing: {
          line: 560,
          lineRule: "exact",
          before: 200,
          after: 200,
        },
      })
    );

    // 4. 渲染问题列表文字
    for (let i = 0; i < problems.length; i++) {
      const problem = problems[i];
      const indexStr = ['一','二','三','四','五','六','七','八','九','十'][i] || (i + 1).toString();
      
      const children = [
        new TextRun({
          text: `${i + 1}. `,
          font: "仿宋_GB2312",
          size: 32,
        })
      ];

      if (problem.department) {
        children.push(
          new TextRun({
            text: problem.department,
            font: "仿宋_GB2312",
            size: 32,
            bold: true,
          })
        );
      } else {
        children.push(
          new TextRun({
            text: "___",
            font: "仿宋_GB2312",
            size: 32,
          })
        );
      }

      children.push(
        new TextRun({
          text: `：${problem.description || '___'}`,
          font: "仿宋_GB2312",
          size: 32,
        })
      );

      if (problem.image) {
        children.push(
          new TextRun({
            text: `（见图${indexStr}）。`,
            font: "仿宋_GB2312",
            size: 32,
          })
        );
      }

      docChildren.push(
        new Paragraph({
          children: children,
          indent: {
            firstLine: 640,
          },
          spacing: {
            line: 560,
            lineRule: "exact",
          },
        })
      );
    }

    // 5. 插入图片
    const problemsWithImages = problems.filter(p => p.image);
    if (problemsWithImages.length > 0) {
      docChildren.push(
        new Paragraph({
          text: "", // 换行
          spacing: { after: 200 },
        })
      );

      // 处理两列排版逻辑（如果只有一张图，直接居中；如果有两张，通过表格或者浮动，这里采用简化的居中堆叠或表格排版）
      // docx 库对于图片的并排排版推荐使用 Table
      // 收集所有带图片的段落
      for (let i = 0; i < problems.length; i++) {
        const problem = problems[i];
        if (!problem.image) continue;
        
        const indexStr = ['一','二','三','四','五','六','七','八','九','十'][i] || (i + 1).toString();
        const dimensions = await getImageDimensions(problem.image);
        
        // 限制最大宽度和高度 (公文页面宽度约 400-500px，设最大宽度为 300)
        const MAX_WIDTH = 300;
        const MAX_HEIGHT = 400;
        let targetWidth = dimensions.width;
        let targetHeight = dimensions.height;
        
        if (targetWidth > MAX_WIDTH) {
          const ratio = MAX_WIDTH / targetWidth;
          targetWidth = MAX_WIDTH;
          targetHeight = targetHeight * ratio;
        }
        if (targetHeight > MAX_HEIGHT) {
          const ratio = MAX_HEIGHT / targetHeight;
          targetHeight = MAX_HEIGHT;
          targetWidth = targetWidth * ratio;
        }

        docChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: base64DataURLToArrayBuffer(problem.image),
                transformation: {
                  width: targetWidth,
                  height: targetHeight,
                },
                type: "png"
              }),
            ],
            spacing: { before: 200 },
          })
        );

        docChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `图${indexStr}`,
                font: "黑体",
                size: 28, // 小四或者五号
                bold: true,
              }),
            ],
            spacing: { after: 400 },
          })
        );
      }
    }

    // 6. 二、改进建议
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "二、改进建议：",
            font: "黑体",
            size: 32,
            bold: true,
          }),
        ],
        spacing: {
          line: 560,
          lineRule: "exact",
          before: 200,
          after: 200,
        },
      })
    );

    for (let i = 0; i < suggestions.length; i++) {
      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${i + 1}. ${suggestions[i].text || '___'}`,
              font: "仿宋_GB2312",
              size: 32,
            }),
          ],
          indent: {
            firstLine: 640,
          },
          spacing: {
            line: 560,
            lineRule: "exact",
          },
        })
      );
    }

    const doc = new Document({
      styles: {
        paragraphStyles: [
          {
            id: "TitleStyle",
            name: "Title Style",
            basedOn: "Normal",
            next: "Normal",
            run: {
              font: "方正小标宋简体",
              size: 44, // 二号 (22pt * 2)
            },
            paragraph: {
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 },
            }
          }
        ]
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: convertInchesToTwip(1.45), // 公文上边距一般为 3.7cm 左右
                bottom: convertInchesToTwip(1.37), // 下边距 3.5cm
                left: convertInchesToTwip(1.1), // 左边距 2.8cm
                right: convertInchesToTwip(1.02), // 右边距 2.6cm
              }
            }
          },
          children: docChildren,
        },
      ],
    });

    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, `${formatDateTitle(inspectionDate)}${inspectionItem}专项检查报告.docx`);
    });
  };

  const formatDateTitle = (dateString: string) => {
    if (!dateString) return '';
    const date = new Date(dateString);
    return `${date.getFullYear()}年${date.getMonth() + 1}月`;
  };

  const formatDateText = (dateString: string) => {
    if (!dateString) return '';
    const date = new Date(dateString);
    return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
  };

  const getDisplayItem = (item: string) => {
    if (item === '麻精药品' || item.includes('毒麻')) {
      return '麻醉、精神、未列管全身麻醉药品';
    }
    return item;
  };

  return (
    <div className="min-h-screen bg-[#f7f7f7] p-4 md:p-8 font-sans">
      <div className="max-w-6xl mx-auto grid grid-cols-1 lg:grid-cols-2 gap-8">
        
        {/* Left Form Section */}
        <div className="space-y-6">
          <div className="flex items-center gap-3 mb-8">
            <div className="w-12 h-12 rounded-2xl bg-duo-green flex items-center justify-center shadow-[0_4px_0_0_var(--color-duo-green-dark)]">
              <FileText className="text-white w-7 h-7" />
            </div>
            <h1 className="text-3xl font-bold text-[#3c3c3c]">生成检查报告</h1>
          </div>

          <div className="card-duo space-y-4">
            <h2 className="text-xl font-bold text-[#3c3c3c]">基本信息</h2>
            
            <div>
              <label className="block text-sm font-bold text-[#afafaf] mb-2 uppercase tracking-wide">检查日期</label>
              <input 
                type="date" 
                value={inspectionDate}
                onChange={e => setInspectionDate(e.target.value)}
                className="input-duo"
              />
            </div>

            <div>
              <label className="block text-sm font-bold text-[#afafaf] mb-2 uppercase tracking-wide">检查科室</label>
              <select 
                value={inspectionDepartment}
                onChange={e => setInspectionDepartment(e.target.value)}
                className="input-duo"
              >
                <option value="内科">内科</option>
                <option value="外科">外科</option>
                <option value="门诊">门诊</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-bold text-[#afafaf] mb-2 uppercase tracking-wide">检查项目</label>
              <select 
                value={inspectionItem}
                onChange={e => setInspectionItem(e.target.value)}
                className="input-duo"
              >
                <option value="高警示药品和抢救车药品">高警示药品和抢救车药品</option>
                <option value="麻精药品">麻精药品</option>
              </select>
            </div>
          </div>

          <div className="card-duo space-y-4">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-[#3c3c3c]">存在的问题</h2>
              <button onClick={handleAddProblem} className="btn-duo btn-duo-blue py-2 px-4 text-sm flex items-center gap-2">
                <Plus size={16} /> 添加问题
              </button>
            </div>

            {problems.map((problem, index) => (
              <div key={problem.id} className="p-4 rounded-xl border-2 border-duo-gray bg-[#f9f9f9] space-y-3 relative group">
                {problems.length > 1 && (
                  <button 
                    onClick={() => handleRemoveProblem(problem.id)}
                    className="absolute -top-3 -right-3 w-8 h-8 bg-duo-red text-white rounded-full flex items-center justify-center shadow-[0_3px_0_0_var(--color-duo-red-dark)] hover:bg-[#ff5e5e] transition-colors"
                  >
                    <Trash2 size={14} />
                  </button>
                )}
                
                <div className="font-bold text-duo-blue">问题 {index + 1}</div>
                
                <input 
                  type="text" 
                  placeholder="问题科室 (如：心内科)"
                  value={problem.department}
                  onChange={e => handleProblemChange(problem.id, 'department', e.target.value)}
                  className="input-duo py-2 text-base"
                />
                
                <textarea 
                  placeholder="存在的问题描述"
                  value={problem.description}
                  onChange={e => handleProblemChange(problem.id, 'description', e.target.value)}
                  className="input-duo py-2 text-base min-h-[80px]"
                />

                <div>
                  <input 
                    type="file" 
                    accept="image/*"
                    onChange={e => handleImageUpload(problem.id, e)}
                    className="hidden"
                    id={`image-upload-${problem.id}`}
                  />
                  <label 
                    htmlFor={`image-upload-${problem.id}`}
                    className="inline-flex items-center gap-2 px-4 py-2 rounded-xl border-2 border-dashed border-duo-blue text-duo-blue font-bold cursor-pointer hover:bg-[#f2fbfd] transition-colors"
                  >
                    <Upload size={18} />
                    {problem.image ? '更换图片' : '上传问题图片'}
                  </label>
                  {problem.image && (
                    <div className="mt-2 rounded-xl overflow-hidden border-2 border-duo-gray max-w-[200px]">
                      <img src={problem.image} alt="问题" className="w-full h-auto" />
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>

          <div className="card-duo space-y-4">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-[#3c3c3c]">改进建议</h2>
              <button onClick={() => handleAddSuggestion()} className="btn-duo btn-duo-blue py-2 px-4 text-sm flex items-center gap-2">
                <Plus size={16} /> 添加建议
              </button>
            </div>

            <div className="bg-[#f9f9f9] p-4 rounded-xl border-2 border-duo-gray mb-4">
              <label className="block text-sm font-bold text-[#afafaf] mb-2 uppercase tracking-wide">建议模板管理</label>
              
              <div className="flex flex-col sm:flex-row gap-2 mb-3">
                <select 
                  className="input-duo py-2 text-base flex-1"
                  value={selectedTemplate}
                  onChange={(e) => setSelectedTemplate(e.target.value)}
                >
                  <option value="" disabled>-- 选择模板 --</option>
                  {suggestionTemplates.map((sug, i) => (
                    <option key={i} value={sug}>{sug}</option>
                  ))}
                </select>
                <button 
                  onClick={() => {
                    if (selectedTemplate) {
                      handleAddSuggestion(selectedTemplate);
                    }
                  }}
                  disabled={!selectedTemplate}
                  className={`btn-duo ${selectedTemplate ? 'btn-duo-blue' : 'bg-duo-gray text-[#afafaf] cursor-not-allowed shadow-none'} py-2 px-4 whitespace-nowrap`}
                >
                  应用
                </button>
                <button 
                  onClick={handleRemoveTemplate}
                  disabled={!selectedTemplate}
                  className={`flex-none w-12 h-12 rounded-xl border-2 flex items-center justify-center transition-colors ${
                    selectedTemplate 
                      ? 'border-duo-red text-duo-red hover:bg-[#fff0f0]' 
                      : 'border-duo-gray text-duo-gray bg-white cursor-not-allowed'
                  }`}
                  title="删除选中模板"
                >
                  <Trash2 size={20} />
                </button>
              </div>

              <div className="flex flex-col sm:flex-row gap-2">
                <input 
                  type="text" 
                  value={newTemplateText}
                  onChange={(e) => setNewTemplateText(e.target.value)}
                  placeholder="输入新的模板内容..."
                  className="input-duo py-2 text-base flex-1"
                />
                <button 
                  onClick={handleAddTemplate}
                  className="btn-duo btn-duo-green py-2 px-4 whitespace-nowrap"
                >
                  保存为新模板
                </button>
              </div>
            </div>

            {suggestions.map((suggestion, index) => (
              <div key={suggestion.id} className="flex gap-2">
                <div className="flex-none w-8 h-8 rounded-full bg-duo-gray text-white flex items-center justify-center font-bold mt-1">
                  {index + 1}
                </div>
                <input 
                  type="text" 
                  value={suggestion.text}
                  onChange={e => handleSuggestionChange(suggestion.id, e.target.value)}
                  className="input-duo py-2 text-base flex-1"
                  placeholder="输入建议内容..."
                />
                {suggestions.length > 1 && (
                  <button 
                    onClick={() => handleRemoveSuggestion(suggestion.id)}
                    className="flex-none w-12 h-12 rounded-xl border-2 border-duo-gray flex items-center justify-center text-duo-red hover:bg-[#fff0f0] hover:border-duo-red transition-colors"
                  >
                    <Trash2 size={20} />
                  </button>
                )}
              </div>
            ))}
          </div>

        </div>

        {/* Right Preview Section */}
        <div className="lg:sticky lg:top-8 h-fit space-y-4">
          <div className="flex justify-between items-center">
            <h2 className="text-2xl font-bold text-[#3c3c3c]">报告预览</h2>
            <button 
              onClick={exportAsWord}
              className="btn-duo btn-duo-green py-2 px-6 flex items-center gap-2"
            >
              <Download size={18} /> 导出Word文档
            </button>
          </div>
          
          <div 
            ref={reportRef}
            className="bg-white p-8 rounded-2xl border-2 border-duo-gray shadow-sm text-[#2a2a2a] leading-relaxed break-words"
            style={{ fontFamily: 'SimSun, "宋体", serif' }}
          >
            <h1 className="text-2xl font-bold text-center mb-6">
              {formatDateTitle(inspectionDate)}{inspectionItem}专项检查报告
            </h1>
            
            <p className="mb-6 text-lg indent-8">
              {formatDateText(inspectionDate)}，药学部药品质量管理工作小组对{inspectionDepartment}科室{getDisplayItem(inspectionItem)}进行检查，现将存在问题整理汇报如下：
            </p>

            <div className="mb-6">
              <h2 className="text-xl font-bold mb-3">一、存在问题：</h2>
              <div className="space-y-2 pl-2 mb-6">
                {problems.map((problem, index) => (
                  <div key={problem.id} className="text-lg">
                    <p>
                      {index + 1}. {problem.department ? <span className="font-bold">{problem.department}</span> : '___'}：
                      {problem.description || '___'}
                      {problem.image && `（见图${['一','二','三','四','五','六','七','八','九','十'][index] || index + 1}）。`}
                    </p>
                  </div>
                ))}
              </div>
              
              {/* 图片区域 - 最多两列 */}
              {problems.some(p => p.image) && (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-6">
                  {problems.map((problem, index) => {
                    if (!problem.image) return null;
                    return (
                      <div key={`img-${problem.id}`} className="flex flex-col items-center border border-gray-200 p-2 rounded bg-gray-50">
                        <div className="w-full h-48 flex items-center justify-center bg-white rounded overflow-hidden">
                          <img 
                            src={problem.image} 
                            alt={`图${['一','二','三','四','五','六','七','八','九','十'][index] || index + 1}`} 
                            className="max-w-full max-h-full object-contain"
                          />
                        </div>
                        <p className="text-sm text-gray-700 mt-2 font-bold">
                          图{['一','二','三','四','五','六','七','八','九','十'][index] || index + 1}
                        </p>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div>
              <h2 className="text-xl font-bold mb-3">二、改进建议：</h2>
              <div className="space-y-2 pl-2">
                {suggestions.map((suggestion, index) => (
                  <div key={suggestion.id} className="text-lg">
                    {index + 1}. {suggestion.text || '___'}
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

      </div>
    </div>
  );
};

export default App;