import React, { useState, useRef, useEffect } from 'react';
import { 
  Bot, 
  Wrench, 
  Play, 
  Trash2, 
  Terminal, 
  Plus, 
  X,
  Server,
  Cloud,
  MessageSquare,
  Send,
  FileSpreadsheet,
  FileText,
  Calculator
} from 'lucide-react';
// Use ESM build of xlsx to avoid bundler resolution issues in Vite
import * as XLSX from 'xlsx/xlsx.mjs';
import * as pdfjsLib from 'pdfjs-dist';
import mammoth from 'mammoth';
import { evaluate } from 'mathjs';

// Configure pdfjs worker
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.js', import.meta.url).toString();

// --- File Parsing Helpers ---
const readFileAsArrayBuffer = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result);
  reader.onerror = reject;
  reader.readAsArrayBuffer(file);
});

const readFileAsText = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result);
  reader.onerror = reject;
  reader.readAsText(file);
});

// --- Helpers: URL extraction ---
const extractUrls = (text) => {
  if (!text) return [];
  const urlRegex = /https?:\/\/[^\s)"']+/g;
  const matches = text.match(urlRegex) || [];
  // dedupe
  return Array.from(new Set(matches));
};

const parseExcelFile = async (file) => {
  const buffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  const rows = json;
  let output = `Excel Sheet: ${sheetName} (${rows.length} rows)\n\n`;
  if (rows.length) {
    output += `Headers: ${rows[0].join(' | ')}\n\n`;
  }
  const maxRows = Math.min(50, rows.length);
  for (let i = 1; i < maxRows; i++) {
    output += `Row ${i}: ${rows[i].join(' | ')}\n`;
  }
  if (rows.length > maxRows) {
    output += `\n... and ${rows.length - maxRows} more rows`;
  }
  return output;
};

const parseCsvOrText = async (file) => {
  const text = await readFileAsText(file);
  return `Text/CSV Content (first 5000 chars):\n${text.slice(0, 5000)}`;
};

const parsePdfFile = async (file) => {
  const buffer = await readFileAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let text = '';
  const maxPages = Math.min(20, pdf.numPages);
  for (let i = 1; i <= maxPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(' ') + '\n\n';
  }
  return `PDF Text (first ${maxPages} pages):\n${text.slice(0, 8000)}`;
};

const parseDocxFile = async (file) => {
  const buffer = await readFileAsArrayBuffer(file);
  const { value } = await mammoth.extractRawText({ arrayBuffer: buffer });
  return `Word Text (first 8000 chars):\n${value.slice(0, 8000)}`;
};

// --- 1. Custom Graph Engine Utilities ---

// Simple Bezier curve generator for smooth connections
const generateBezierPath = (start, end) => {
  const controlPointX = start.x + (end.x - start.x) / 2;
  return `M ${start.x} ${start.y} C ${controlPointX} ${start.y} ${controlPointX} ${end.y} ${end.x} ${end.y}`;
};

// Node Component (The draggable box)
const Node = ({ id, type, x, y, data, isSelected, onMouseDown, onHandleMouseDown, onDelete, onUpdate }) => {
  const isAgent = type === 'agent';
  const isTool = type === 'tool';
  const isTrigger = type === 'trigger';

  // Local state for inputs to prevent re-rendering entire graph on every keystroke
  const handleInputChange = (field, value) => {
    onUpdate(id, { ...data, [field]: value });
  };

  return (
    <div 
      className={`absolute shadow-xl rounded-xl flex flex-col transition-shadow ${isSelected ? 'ring-2 ring-blue-500 shadow-blue-500/20' : ''} ${isAgent ? 'w-80 bg-slate-900 border-2 border-slate-700' : isTool ? 'w-56 bg-slate-900 border border-orange-500/30' : 'w-auto bg-slate-900 border border-green-500/50 rounded-full'}`}
      style={{ left: x, top: y, cursor: 'grab' }}
      onMouseDown={(e) => onMouseDown(e, id)}
    >
      {/* Handles (Inputs/Outputs) */}
      <div className="absolute inset-0 pointer-events-none">
        {/* Input Handle (Left) */}
        {!isTrigger && (
          <div 
            className="absolute left-0 top-1/2 -translate-x-1/2 -translate-y-1/2 w-3 h-3 bg-blue-500 border border-slate-900 rounded-full pointer-events-auto cursor-crosshair hover:scale-125 transition-transform"
            onMouseDown={(e) => onHandleMouseDown(e, id, 'in')}
            title="Input"
          />
        )}
        
        {/* Output Handle (Right) */}
        {!isTool && (
          <div 
            className="absolute right-0 top-1/2 -translate-x-1/2 -translate-y-1/2 w-3 h-3 bg-green-500 border border-slate-900 rounded-full pointer-events-auto cursor-crosshair hover:scale-125 transition-transform"
            onMouseDown={(e) => onHandleMouseDown(e, id, 'out')}
            title="Output"
          />
        )}

        {/* Tools Handle (Bottom) - Only for Agents */}
        {isAgent && (
          <div 
            className="absolute bottom-0 left-1/2 -translate-x-1/2 translate-y-1/2 w-3 h-3 bg-orange-500 border border-slate-900 rounded-sm pointer-events-auto cursor-crosshair hover:scale-125 transition-transform"
            onMouseDown={(e) => onHandleMouseDown(e, id, 'tools')}
            title="Tools"
          />
        )}

        {/* Tool Output (Top) - Only for Tools */}
        {isTool && (
          <div 
            className="absolute top-0 left-1/2 -translate-x-1/2 -translate-y-1/2 w-2 h-2 bg-orange-500 border border-slate-900 rounded-sm pointer-events-auto cursor-crosshair hover:scale-125 transition-transform"
            onMouseDown={(e) => onHandleMouseDown(e, id, 'tool-out')}
            title="Tool Output"
          />
        )}
      </div>

      {/* Node Content */}
      {isAgent && (
        <>
          <div className="bg-slate-800 p-3 flex items-center justify-between border-b border-slate-700 rounded-t-xl">
            <div className="flex items-center gap-2">
              <div className="bg-blue-500/20 p-1.5 rounded-lg">
                <Bot size={16} className="text-blue-400" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-100">AI Agent</h3>
                <p className="text-[10px] text-slate-400">LLM Processor</p>
              </div>
            </div>
            <button onClick={(e) => { e.stopPropagation(); onDelete(id); }} className="text-slate-500 hover:text-red-400"><X size={14}/></button>
          </div>
          <div className="p-3 space-y-3 bg-slate-900/50 rounded-b-xl">
            <div className="flex bg-slate-950 rounded p-1 gap-1">
              <button 
                className={`flex-1 text-[10px] py-1 rounded flex items-center justify-center gap-1 ${!data.isLocal ? 'bg-blue-600 text-white' : 'text-slate-500 hover:text-slate-300'}`}
                onClick={(e) => { e.stopPropagation(); handleInputChange('isLocal', false); }}
              >
                <Cloud size={10} /> Cloud
              </button>
              <button 
                className={`flex-1 text-[10px] py-1 rounded flex items-center justify-center gap-1 ${data.isLocal ? 'bg-blue-600 text-white' : 'text-slate-500 hover:text-slate-300'}`}
                onClick={(e) => { e.stopPropagation(); handleInputChange('isLocal', true); }}
              >
                <Server size={10} /> Local
              </button>
            </div>

            <textarea 
              className="w-full bg-slate-950 text-xs text-slate-300 p-2 rounded border border-slate-800 focus:border-blue-500 outline-none resize-none h-16"
              defaultValue={data.systemPrompt}
              placeholder="System Prompt..."
              onChange={(e) => handleInputChange('systemPrompt', e.target.value)}
              onMouseDown={(e) => e.stopPropagation()} 
            />

            {data.isLocal ? (
               <div className="space-y-2 animate-in fade-in slide-in-from-top-1 duration-200">
                  <div>
                    <label className="text-[9px] uppercase text-slate-500 font-bold mb-1 block">Local Endpoint (LM Studio/Ollama)</label>
                    <input 
                      className="w-full bg-slate-950 text-xs text-slate-300 p-2 rounded border border-slate-800 focus:border-blue-500 outline-none font-mono"
                      defaultValue={data.baseUrl || "http://172.16.0.2:1234/v1"}
                      placeholder="http://172.16.0.2:1234/v1 (LM Studio) or http://localhost:11434 (Ollama)"
                      onChange={(e) => handleInputChange('baseUrl', e.target.value)}
                      onMouseDown={(e) => e.stopPropagation()}
                    />
                    <p className="text-[8px] text-slate-600 mt-1">LM Studio: Add /v1 at the end (e.g., http://172.16.0.2:1234/v1)</p>
                  </div>
                  <div>
                    <label className="text-[9px] uppercase text-slate-500 font-bold mb-1 block">Model Name</label>
                    <input 
                      className="w-full bg-slate-950 text-xs text-slate-300 p-2 rounded border border-slate-800 focus:border-blue-500 outline-none font-mono"
                      defaultValue={data.modelName || "llama3"}
                      placeholder="e.g. llama3, mistral"
                      onChange={(e) => handleInputChange('modelName', e.target.value)}
                      onMouseDown={(e) => e.stopPropagation()}
                    />
                  </div>
               </div>
            ) : (
               <select 
                 className="w-full bg-slate-950 text-xs text-slate-300 p-2 rounded border border-slate-800 outline-none"
                 defaultValue={data.model}
                 onChange={(e) => handleInputChange('model', e.target.value)}
                 onMouseDown={(e) => e.stopPropagation()}
               >
                 <option value="gpt-4o">GPT-4o</option>
                 <option value="claude-3-5">Claude 3.5 Sonnet</option>
                 <option value="gemini-pro">Gemini Pro</option>
               </select>
            )}
          </div>
        </>
      )}

      {isTool && (
        <div className="rounded-lg overflow-hidden">
          <div className="bg-orange-500/10 p-2 flex items-center justify-between border-b border-orange-500/20">
            <div className="flex items-center gap-2">
              {data.toolType === 'google-sheets' ? (
                <FileSpreadsheet size={14} className="text-orange-400" />
              ) : data.toolType === 'local-file' ? (
                <FileText size={14} className="text-orange-400" />
              ) : data.toolType === 'calculator' ? (
                <Calculator size={14} className="text-orange-400" />
              ) : (
                <Wrench size={14} className="text-orange-400" />
              )}
              <span className="text-xs font-bold text-orange-100">{data.label}</span>
            </div>
            <button onClick={(e) => { e.stopPropagation(); onDelete(id); }} className="text-orange-400/50 hover:text-orange-400"><X size={12}/></button>
          </div>
          <div className="p-2 bg-slate-900 space-y-2">
            <div className="text-[10px] text-slate-400">{data.description}</div>
            {data.toolType === 'google-sheets' && (
              <div className="space-y-1 text-[9px]">
                <div className="text-slate-500">
                  <span className="font-bold">Sheet ID:</span> {data.sheetId ? data.sheetId.substring(0, 20) + '...' : 'Not set'}
                </div>
                <div className="text-slate-500">
                  <span className="font-bold">Range:</span> {data.range || 'A1:Z1000'}
                </div>
              </div>
            )}
            {data.toolType === 'local-file' && (
              <div className="space-y-1 text-[9px]">
                <div className="text-slate-500">
                  <span className="font-bold">File:</span> {data.fileName || 'Not uploaded'}
                </div>
                <div className="text-slate-500">
                  <span className="font-bold">Type:</span> {data.fileType || 'auto'}
                </div>
              </div>
            )}
          </div>
          {data.toolType === 'google-sheets' && (
            <div className="p-2 bg-slate-950 border-t border-orange-500/20 space-y-2">
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">Google Sheets ID</label>
                <input 
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none font-mono"
                  defaultValue={data.sheetId || ""}
                  placeholder="Enter Google Sheets ID"
                  onChange={(e) => handleInputChange('sheetId', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                />
              </div>
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">Range (e.g., A1:Z1000)</label>
                <input 
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none font-mono"
                  defaultValue={data.range || "A1:Z1000"}
                  placeholder="A1:Z1000"
                  onChange={(e) => handleInputChange('range', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                />
              </div>
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">API Key (Optional - for private sheets)</label>
                <input 
                  type="password"
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none font-mono"
                  defaultValue={data.apiKey || ""}
                  placeholder="Google API Key (optional)"
                  onChange={(e) => handleInputChange('apiKey', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                />
              </div>
            </div>
          )}
          {data.toolType === 'web-search' && (
            <div className="p-2 bg-slate-950 border-t border-orange-500/20 space-y-2">
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">Provider</label>
                <select
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none"
                  defaultValue={data.provider || 'bing'}
                  onChange={(e) => handleInputChange('provider', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  <option value="bing">Bing</option>
                  <option value="google">Google (requires custom API)</option>
                </select>
              </div>
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">API Key (Optional - required for programmatic Bing search)</label>
                <input 
                  type="password"
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none font-mono"
                  defaultValue={data.apiKey || ""}
                  placeholder="Bing API Key (optional)"
                  onChange={(e) => handleInputChange('apiKey', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                />
              </div>
            </div>
          )}
          {data.toolType === 'local-file' && (
            <div className="p-2 bg-slate-950 border-t border-orange-500/20 space-y-2">
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">File Type</label>
                <select
                  className="w-full bg-slate-900 text-[10px] text-slate-300 p-1.5 rounded border border-slate-800 focus:border-orange-500 outline-none"
                  defaultValue={data.fileType || "auto"}
                  onChange={(e) => handleInputChange('fileType', e.target.value)}
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  <option value="auto">Auto Detect</option>
                  <option value="excel">Excel (.xlsx/.xls)</option>
                  <option value="csv">CSV</option>
                  <option value="pdf">PDF</option>
                  <option value="docx">Word (.docx)</option>
                  <option value="txt">Text</option>
                </select>
              </div>
              <div>
                <label className="text-[8px] uppercase text-slate-500 font-bold mb-1 block">Upload File</label>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv,.pdf,.docx,.txt"
                  className="w-full text-[10px] text-slate-300"
                  onChange={async (e) => {
                    e.stopPropagation();
                    const file = e.target.files?.[0];
                    if (!file) return;
                    try {
                      const inferredType = data.fileType === 'auto'
                        ? (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) ? 'excel'
                          : file.name.toLowerCase().endsWith('.csv') ? 'csv'
                          : file.name.toLowerCase().endsWith('.pdf') ? 'pdf'
                          : file.name.toLowerCase().endsWith('.docx') ? 'docx'
                          : 'txt'
                        : data.fileType;
                      onUpdate(id, { ...data, loading: true, fileName: file.name, fileType: inferredType });
                      let parsed = '';
                      if (inferredType === 'excel') parsed = await parseExcelFile(file);
                      else if (inferredType === 'csv' || inferredType === 'txt') parsed = await parseCsvOrText(file);
                      else if (inferredType === 'pdf') parsed = await parsePdfFile(file);
                      else if (inferredType === 'docx') parsed = await parseDocxFile(file);
                      else parsed = await parseCsvOrText(file);
                      onUpdate(id, { ...data, fileContent: parsed, fileName: file.name, fileType: inferredType, loading: false });
                    } catch (err) {
                      onUpdate(id, { ...data, loading: false, fileContent: `Error parsing file: ${err.message}` });
                    }
                  }}
                />
              </div>
              {data.loading && <div className="text-[10px] text-orange-400">Parsing file...</div>}
              {data.fileContent && (
                <div className="bg-slate-900 text-[9px] text-slate-300 p-2 rounded border border-slate-800 max-h-32 overflow-auto">
                  {data.fileContent.slice(0, 500)}{data.fileContent.length > 500 ? ' ...' : ''}
                </div>
              )}
            </div>
          )}
          {data.toolType === 'calculator' && (
            <div className="p-2 bg-slate-950 border-t border-orange-500/20 text-[10px] text-slate-400">
              This calculator tool evaluates mathematical expressions (powered by mathjs). The agent can send expressions to this tool to compute results (supports +, -, *, /, parentheses, functions like sqrt, sin, etc.).
            </div>
          )}
        </div>
      )}

      {isTrigger && (
        <div className="px-4 py-2 flex items-center gap-3">
          <div className="bg-green-500/20 p-1.5 rounded-full">
            <Play size={14} className="text-green-400" />
          </div>
          <div>
            <div className="text-xs font-bold text-green-100">Chat Input</div>
            <div className="text-[9px] text-green-500/70">Starts Flow</div>
          </div>
        </div>
      )}
    </div>
  );
};

// --- 2. Main Application ---

const initialNodes = [
  { id: 'trigger-1', type: 'trigger', x: 100, y: 300, data: { label: 'User Chat' } },
  { id: 'agent-1', type: 'agent', x: 450, y: 150, data: { systemPrompt: 'You are a helpful assistant that can answer questions using uploaded files (Excel, Word, PDF, CSV, TXT), Google Sheets data, and calculator results. Use all provided context to answer accurately.', isLocal: true, baseUrl: 'http://172.16.0.2:1234/v1', modelName: 'llama3' } },
  { id: 'tool-1', type: 'tool', x: 450, y: 520, data: { label: 'Google Sheets', description: 'Read data from Google Sheets', toolType: 'google-sheets', sheetId: '', range: 'A1:Z1000', apiKey: '' } },
  { id: 'tool-2', type: 'tool', x: 200, y: 520, data: { label: 'Local File', description: 'Upload and read Excel / Word / PDF / CSV / TXT files', toolType: 'local-file', fileType: 'auto', fileName: '', fileContent: '', loading: false } },
  { id: 'tool-3', type: 'tool', x: 700, y: 520, data: { label: 'Calculator', description: 'Custom calculator for expressions', toolType: 'calculator' } },
];

const initialEdges = [
  { id: 'e1', source: 'trigger-1', target: 'agent-1', type: 'flow' },
  { id: 'e2', source: 'tool-1', target: 'agent-1', type: 'tool' }, 
];

export default function AgentBuilderApp() {
  const [nodes, setNodes] = useState(initialNodes);
  const [edges, setEdges] = useState(initialEdges);
  const [selectedNode, setSelectedNode] = useState(null);
  
  // Canvas State
  const [offset, setOffset] = useState({ x: 0, y: 0 });
  const [zoom, setZoom] = useState(1);
  const [isDraggingCanvas, setIsDraggingCanvas] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  
  // Node Dragging State
  const [isDraggingNode, setIsDraggingNode] = useState(false);
  const [draggedNodeId, setDraggedNodeId] = useState(null);
  const [nodeDragOffset, setNodeDragOffset] = useState({ x: 0, y: 0 });

  // Connection State
  const [connecting, setConnecting] = useState(null); 
  const [mousePos, setMousePos] = useState({ x: 0, y: 0 });

  // Execution State
  const [logs, setLogs] = useState([]);
  const [isRunning, setIsRunning] = useState(false);
  
  // Chat State
  const [messages, setMessages] = useState([]);
  const [inputMessage, setInputMessage] = useState('');
  const [showChat, setShowChat] = useState(true);
  const chatEndRef = useRef(null);
  
  // Scroll to bottom of chat
  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // --- Helpers ---
  const screenToWorld = (sx, sy) => ({
    x: (sx - offset.x) / zoom,
    y: (sy - offset.y) / zoom
  });

  // --- Event Handlers ---

  const handleMouseDown = (e) => {
    if (e.target.id === 'canvas-bg') {
      setIsDraggingCanvas(true);
      setDragStart({ x: e.clientX - offset.x, y: e.clientY - offset.y });
      setSelectedNode(null);
    }
  };

  const handleNodeUpdate = (id, newData) => {
    setNodes(prev => prev.map(n => n.id === id ? { ...n, data: newData } : n));
  };

  const handleNodeMouseDown = (e, id) => {
    e.stopPropagation();
    setSelectedNode(id);
    setIsDraggingNode(true);
    setDraggedNodeId(id);
    const node = nodes.find(n => n.id === id);
    const worldMouse = screenToWorld(e.clientX, e.clientY);
    setNodeDragOffset({ x: worldMouse.x - node.x, y: worldMouse.y - node.y });
  };

  const handleHandleMouseDown = (e, nodeId, handleType) => {
    e.stopPropagation();
    const node = nodes.find(n => n.id === nodeId);
    
    // Calculate exact handle position based on node geometry (simplified)
    let startX = node.x;
    let startY = node.y;

    const isAgent = node.type === 'agent';
    const isTool = node.type === 'tool';
    
    // Width/Height assumptions (matching Tailwind classes roughly)
    const w = isAgent ? 320 : isTool ? 224 : 150; 
    const h = isAgent ? 280 : isTool ? 100 : 60; // Approximate heights

    if (handleType === 'out') { startX += w; startY += h/2; }
    if (handleType === 'in') { startY += h/2; }
    if (handleType === 'tools') { startX += w/2; startY += h; }
    if (handleType === 'tool-out') { startX += w/2; }

    setConnecting({ nodeId, handleType, startX, startY });
  };

  const handleMouseMove = (e) => {
    setMousePos({ x: e.clientX, y: e.clientY });

    if (isDraggingCanvas) {
      setOffset({ x: e.clientX - dragStart.x, y: e.clientY - dragStart.y });
    }

    if (isDraggingNode && draggedNodeId) {
      const worldMouse = screenToWorld(e.clientX, e.clientY);
      setNodes(prev => prev.map(n => {
        if (n.id === draggedNodeId) {
          return { ...n, x: worldMouse.x - nodeDragOffset.x, y: worldMouse.y - nodeDragOffset.y };
        }
        return n;
      }));
    }
  };

  const handleMouseUp = (e) => {
    setIsDraggingCanvas(false);
    setIsDraggingNode(false);
    setDraggedNodeId(null);

    // Finalize Connection
    if (connecting) {
      const worldMouse = screenToWorld(e.clientX, e.clientY);
      const targetNode = nodes.find(n => 
        worldMouse.x > n.x && worldMouse.x < n.x + 200 &&
        worldMouse.y > n.y && worldMouse.y < n.y + 100 &&
        n.id !== connecting.nodeId
      );

      if (targetNode) {
        let type = 'flow';
        let source = connecting.nodeId;
        let target = targetNode.id;

        // For tool connections, ensure the tool is the source and agent is the target
        if (connecting.handleType === 'tools' || connecting.handleType === 'tool-out') {
          type = 'tool';
          // If dragging from agent's tools handle, swap so tool is source
          if (connecting.handleType === 'tools') {
            source = targetNode.id;
            target = connecting.nodeId;
          }
        }

        const newEdge = {
          id: `e-${Date.now()}`,
          source,
          target,
          type
        };
        setEdges(prev => [...prev, newEdge]);
      }
      setConnecting(null);
    }
  };

  const handleDeleteNode = (id) => {
    setNodes(prev => prev.filter(n => n.id !== id));
    setEdges(prev => prev.filter(e => e.source !== id && e.target !== id));
  };

  const addLog = (type, message) => {
    setLogs(prev => [...prev, { type, message, timestamp: new Date().toLocaleTimeString() }]);
  };
  
  const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  // --- Local LLM API Call ---
  const callLocalLLM = async (agent, userMessage, conversationHistory = []) => {
    let baseUrl = agent.data.baseUrl || 'http://172.16.0.2:1234/v1';
    const modelName = agent.data.modelName || 'llama3';
    const systemPrompt = agent.data.systemPrompt || 'You are a helpful assistant.';

    try {
      // Check if using OpenAI-compatible endpoint (contains /v1) - LM Studio uses this
      // Also check if it's a common LM Studio port (1234) and auto-add /v1 if missing
      const isLMStudioPort = baseUrl.includes(':1234') && !baseUrl.includes('/v1');
      const isOpenAICompatible = baseUrl.includes('/v1') || isLMStudioPort;
      
      if (isOpenAICompatible) {
        // OpenAI-compatible API (LM Studio, Ollama with OpenAI compatibility)
        const messages = [
          { role: 'system', content: systemPrompt },
          ...conversationHistory,
          { role: 'user', content: userMessage }
        ];

        // Ensure baseUrl ends with /v1 for LM Studio
        let apiUrl = baseUrl;
        if (isLMStudioPort || (!baseUrl.endsWith('/v1') && baseUrl.includes(':1234'))) {
          apiUrl = baseUrl.endsWith('/') ? `${baseUrl}v1` : `${baseUrl}/v1`;
        }
        
        addLog("system", `Calling LM Studio API: ${apiUrl}/chat/completions`);

        const response = await fetch(`${apiUrl}/chat/completions`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            model: modelName,
            messages: messages,
            stream: false,
            temperature: 0.7
          })
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
        }

        const data = await response.json();
        return data.choices[0]?.message?.content || 'No response generated.';
      } else {
        // Ollama Native API
        // Normalize baseUrl - remove trailing /v1 if present
        const normalizedUrl = baseUrl.replace(/\/v1\/?$/, '');
        
        // Convert conversation history to Ollama format
        const ollamaMessages = [];
        
        // Add system prompt as first message if provided
        if (systemPrompt) {
          ollamaMessages.push({ role: 'system', content: systemPrompt });
        }
        
        // Add conversation history
        conversationHistory.forEach(msg => {
          ollamaMessages.push({ role: msg.role, content: msg.content });
        });
        
        // Add current user message
        ollamaMessages.push({ role: 'user', content: userMessage });

        addLog("system", `Calling Ollama API: ${normalizedUrl}/api/chat`);

        const response = await fetch(`${normalizedUrl}/api/chat`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            model: modelName,
            messages: ollamaMessages,
            stream: false
          })
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
        }

        const data = await response.json();
        return data.message?.content || 'No response generated.';
      }
    } catch (error) {
      console.error('Error calling local LLM:', error);
      throw error;
    }
  };

  // --- Google Sheets Data Fetching ---
  const fetchGoogleSheetsData = async (sheetId, range = 'A1:Z1000', apiKey = null) => {
    try {
      // Construct the Google Sheets API URL
      // For public sheets, no API key needed
      // For private sheets, API key is required
      let url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;
      
      if (apiKey) {
        url += `?key=${apiKey}`;
      } else {
        // For public sheets, add alt=json parameter
        url += '?alt=json';
      }

      addLog("tool", `Fetching data from Google Sheets: ${sheetId}`);
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to fetch Google Sheets data: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      
      if (data.values && data.values.length > 0) {
        // Format the data as a readable table
        const rows = data.values;
        let formattedData = `Google Sheets Data (${rows.length} rows):\n\n`;
        
        // Use first row as headers if available
        const headers = rows[0];
        formattedData += `Headers: ${headers.join(' | ')}\n\n`;
        
        // Add data rows (limit to first 50 rows for context)
        const maxRows = Math.min(50, rows.length);
        for (let i = 1; i < maxRows; i++) {
          formattedData += `Row ${i}: ${rows[i].join(' | ')}\n`;
        }
        
        if (rows.length > maxRows) {
          formattedData += `\n... and ${rows.length - maxRows} more rows`;
        }
        
        addLog("success", `Successfully fetched ${rows.length} rows from Google Sheets`);
        return formattedData;
      } else {
        return "Google Sheets is empty or range has no data.";
      }
    } catch (error) {
      console.error('Error fetching Google Sheets:', error);
      addLog("error", `Google Sheets error: ${error.message}`);
      throw error;
    }
  };

  // --- Tool Execution ---
  const executeTool = async (tool, input) => {
    await wait(300);
    // Web Search (Bing) - Always fetch and read page content via text-proxy, no API key needed
    if (tool.data.toolType === 'web-search') {
      const q = (input || '').trim();
      if (!q) return 'No query provided for web search.';

      // Helper to fetch with timeout to prevent slow fetches from blocking
      const fetchWithTimeout = (url, timeoutMs = 5000) => {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeoutMs);
        return fetch(url, { signal: controller.signal })
          .finally(() => clearTimeout(timer));
      };

      // Always use text-proxy to fetch search results and top page content (no API key needed)
      try {
        addLog('tool', `Searching for: "${q}"`);
        const searchProxy = `https://r.jina.ai/http://www.bing.com/search?q=${encodeURIComponent(q)}`;
        const sresp = await fetchWithTimeout(searchProxy, 8000);
        if (sresp.ok) {
          const stext = await sresp.text();
          // Extract all URLs from search results
          const urlMatches = stext.match(/https?:\/\/[^\s)"'<>]+/gi) || [];
          const urls = [...new Set(urlMatches)].filter(u => !u.includes('bing.com')).slice(0, 5);
          
          if (urls.length === 0) {
            return `No results found for: "${q}"`;
          }

          const topUrl = urls[0];
          addLog('tool', `Found top URL: ${topUrl}`);
          
          // Fetch and read the top result page
          try {
            const jinaUrl = 'https://r.jina.ai/http://' + topUrl.replace(/^https?:\/\//i, '');
            const pageResp = await fetchWithTimeout(jinaUrl, 8000);
            if (pageResp.ok) {
              let pageText = await pageResp.text();
              // Limit to reasonable size for fast processing
              pageText = pageText.slice(0, 2000);
              addLog('success', `Fetched content from ${topUrl}`);
              return `URL: ${topUrl}\n\nContent Summary:\n${pageText}`;
            }
          } catch (err) {
            addLog('warn', `Could not fetch page content: ${err.message}`);
            return `URL: ${topUrl}\n\nCould not retrieve full content, but this is the top result for "${q}"`;
          }
        }
        return `Could not perform search. Please try again.`;
      } catch (err) {
        addLog('error', `Search error: ${err.message}`);
        return `Search failed: ${err.message}`;
      }
    }
    
    // === Legacy: API-based Bing search (kept for backward compatibility) ===
    if (tool.data.toolType === 'web-search-api' && tool.data.apiKey) {
      const provider = tool.data.provider || 'bing';
      const apiKey = tool.data.apiKey;
      const q = (input || '').trim();
      if (!q) return 'No query provided for web search.';

      // Helper to fetch with timeout to prevent slow fetches from blocking
      const fetchWithTimeout = (url, timeoutMs = 5000) => {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeoutMs);
        return fetch(url, { signal: controller.signal })
          .finally(() => clearTimeout(timer));
      };

      if (provider === 'bing' && apiKey) {
        try {
          const resp = await fetch(`https://api.bing.microsoft.com/v7.0/search?q=${encodeURIComponent(q)}&count=3`, {
            headers: { 'Ocp-Apim-Subscription-Key': apiKey }
          });
          if (!resp.ok) {
            const txt = await resp.text();
            throw new Error(`Bing search failed: ${resp.status} ${txt}`);
          }
          const json = await resp.json();
          let out = `Bing Search Results for "${q}":\n\n`;
          if (json.webPages && json.webPages.value && json.webPages.value.length) {
            const results = json.webPages.value.slice(0,3);
            results.forEach((v, i) => {
              out += `${i+1}. ${v.name} - ${v.url}\n${v.snippet || ''}\n\n`;
            });

            // Try to fetch the top result page and extract text for richer context (may fail due to CORS)
            const top = results[0];
            try {
              addLog('tool', `Fetching top result for richer context: ${top.url}`);
              const pageResp = await fetchWithTimeout(top.url, 5000);
              if (pageResp.ok) {
                const html = await pageResp.text();
                // Strip HTML tags to get plain text
                const text = html.replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, ' ')
                                  .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, ' ')
                                  .replace(/<[^>]+>/g, ' ')
                                  .replace(/\s+/g, ' ')
                                  .trim();
                const short = text.slice(0, 3000);
                out += `Top result content (first ${short.length} chars):\n${short}\n\n`;
              } else {
                out += `Could not fetch top result (status ${pageResp.status}). Trying text-proxy fallback.\n\n`;
                // fall through to proxy attempt
                throw new Error(`Direct fetch failed with status ${pageResp.status}`);
              }
            } catch (fetchErr) {
              // Try text-proxy fallback (r.jina.ai) to bypass CORS and get page text
              try {
                addLog('tool', `Falling back to text-proxy for: ${top.url}`);
                const jinaUrl = 'https://r.jina.ai/http://' + top.url.replace(/^https?:\/\//i, '');
                const proxyResp = await fetchWithTimeout(jinaUrl, 5000);
                if (proxyResp.ok) {
                  const proxyText = await proxyResp.text();
                  const short = proxyText.slice(0, 3000);
                  out += `Top result content (via text-proxy, first ${short.length} chars):\n${short}\n\n`;
                } else {
                  out += `Top result could not be fetched via proxy (status ${proxyResp.status}). Using snippet only.\n\n`;
                }
              } catch (proxyErr) {
                addLog('warn', `Proxy fetch failed: ${proxyErr.message}`);
                out += `Top result could not be fetched due to CORS or network. Using snippet only.\n\n`;
              }
            }
          } else {
            out += 'No results found.';
          }
          return out;
        } catch (err) {
          addLog('error', `Web search error: ${err.message}`);
          return `Web search error: ${err.message}`;
        }
      } else {
        // No API key: attempt a text-proxy search fetch so we can extract the top result and read it (works around CORS)
        try {
          addLog('tool', 'No API key provided — using text-proxy search fallback');
          const searchProxy = `https://r.jina.ai/http://www.bing.com/search?q=${encodeURIComponent(q)}`;
          const sresp = await fetch(searchProxy);
          if (sresp.ok) {
            const stext = await sresp.text();
            // Try to extract first URL from search results text
            const urlMatch = stext.match(/https?:\/\/[^\s)"']+/i);
            if (urlMatch && urlMatch[0]) {
              const topUrl = urlMatch[0];
              addLog('tool', `Found top URL via proxy search: ${topUrl}`);
              try {
                const jinaUrl = 'https://r.jina.ai/http://' + topUrl.replace(/^https?:\/\//i, '');
                const pageResp = await fetch(jinaUrl);
                if (pageResp.ok) {
                  const pageText = await pageResp.text();
                  const snippet = pageText.slice(0, 8000);
                  return `Top result: ${topUrl}\n\n${snippet}`;
                }
              } catch (err) {
                addLog('warn', `Proxy fetch of topUrl failed: ${err.message}`);
              }
            }
          }
        } catch (err) {
          addLog('warn', `Proxy search failed: ${err.message}`);
        }

        // Fallback to returning the search URL when proxy approach fails
        const url = `https://www.bing.com/search?q=${encodeURIComponent(q)}`;
        return `Search URL: ${url}`;
      }
    }
    
    // Google Sheets
    if (tool.data.toolType === 'google-sheets') {
      if (!tool.data.sheetId) {
        throw new Error('Google Sheets ID is required. Please configure the tool.');
      }
      const sheetId = tool.data.sheetId;
      const range = tool.data.range || 'A1:Z1000';
      const apiKey = tool.data.apiKey || null;
      return await fetchGoogleSheetsData(sheetId, range, apiKey);
    }

    // Local files (Excel, CSV, PDF, DOCX, TXT)
    if (tool.data.toolType === 'local-file') {
      if (!tool.data.fileContent) {
        throw new Error('Please upload a file to this tool first.');
      }
      return `File "${tool.data.fileName || 'uploaded file'}" content:\n${tool.data.fileContent}`;
    }

    // Calculator
    if (tool.data.toolType === 'calculator') {
      try {
        const result = evaluate(input || '0');
        return `Calculation result: ${result}`;
      } catch (err) {
        return `Could not calculate: ${err.message}`;
      }
    }
    
    // Mock search
    if (tool.data.label.toLowerCase().includes('search')) {
      return `Search results for "${input}": Found relevant information about the topic.`;
    }
    
    return `Tool "${tool.data.label}" executed with input: ${input}`;
  };

  // --- Real Workflow Execution ---
  const executeWorkflow = async (userMessage) => {
    if (isRunning) return;
    setIsRunning(true);
    
    // Add user message to chat
    const userMsg = { role: 'user', content: userMessage, timestamp: new Date() };
    setMessages(prev => [...prev, userMsg]);
    addLog("info", `User: ${userMessage}`);

    try {
      // Find trigger node
      const trigger = nodes.find(n => n.type === 'trigger');
      if (!trigger) {
        addLog("error", "Error: No Trigger node found.");
        setMessages(prev => [...prev, { 
          role: 'assistant', 
          content: 'Error: No trigger node configured in workflow.', 
          timestamp: new Date() 
        }]);
        setIsRunning(false);
        return;
      }

      addLog("process", `Trigger activated: [${trigger.data.label}]`);

      // Find flow edge from trigger
      const flowEdge = edges.find(e => e.source === trigger.id && e.type === 'flow');
      if (!flowEdge) {
        addLog("warn", "Warning: Trigger not connected to any Agent.");
        setMessages(prev => [...prev, { 
          role: 'assistant', 
          content: 'Error: Trigger is not connected to an agent. Please connect the trigger to an agent node.', 
          timestamp: new Date() 
        }]);
        setIsRunning(false);
        return;
      }

      // Get the agent
      const agent = nodes.find(n => n.id === flowEdge.target);
      if (!agent || agent.type !== 'agent') {
        addLog("error", "Error: Invalid agent node.");
        setIsRunning(false);
        return;
      }

      addLog("process", `Routing to Agent: [${agent.id}]`);

      // Check if local or cloud
      if (agent.data.isLocal) {
        addLog("system", `Connecting to Local API: ${agent.data.baseUrl}`);
        addLog("system", `Using Model: ${agent.data.modelName}`);

        // Get connected tools
        const connectedTools = edges
          .filter(e => e.target === agent.id && e.type === 'tool')
          .map(e => nodes.find(n => n.id === e.source))
          .filter(n => n && n.type === 'tool');

        let toolResults = '';
        if (connectedTools.length > 0) {
          addLog("info", `Agent has ${connectedTools.length} connected tool(s).`);
          
          // Execute tools (simplified - in real scenario, agent would decide which tools to use)
          for (const tool of connectedTools) {
            addLog("tool", `Executing tool: [${tool.data.label}]`);
            const toolResult = await executeTool(tool, userMessage);
            toolResults += `\n[${tool.data.label}]: ${toolResult}\n`;
            addLog("success", `Tool [${tool.data.label}] completed.`);
          }
        }
        
        // Extract URLs from tool results to show as clickable actions in the chat
        const extractedUrls = extractUrls(toolResults);

        // Prepare message with tool results
        const enhancedMessage = toolResults 
          ? `${userMessage}\n\nContext from tools:${toolResults}`
          : userMessage;

        // Get conversation history (last 5 messages for context)
        const recentMessages = messages
          .slice(-5)
          .map(msg => ({ role: msg.role, content: msg.content }));

        // Call local LLM
        addLog("agent", "Calling local LLM...");
        const response = await callLocalLLM(agent, enhancedMessage, recentMessages);
        
        addLog("success", "Response generated successfully.");
        
        // Add assistant response to chat (include any extracted URLs so UI can render action buttons)
        setMessages(prev => [...prev, { 
          role: 'assistant', 
          content: response, 
          timestamp: new Date(),
          agentId: agent.id,
          urls: extractedUrls
        }]);
      } else {
        addLog("system", "Cloud API not implemented yet. Please use Local mode.");
        setMessages(prev => [...prev, { 
          role: 'assistant', 
          content: 'Cloud API mode is not yet implemented. Please configure the agent to use Local mode.', 
          timestamp: new Date() 
        }]);
      }
    } catch (error) {
      addLog("error", `Error: ${error.message}`);
      setMessages(prev => [...prev, { 
        role: 'assistant', 
        content: `Error: ${error.message}. Make sure your local LLM server (Ollama/LM Studio) is running.`, 
        timestamp: new Date() 
      }]);
    } finally {
      setIsRunning(false);
    }
  };

  // --- Handle Chat Submit ---
  const handleChatSubmit = (e) => {
    e.preventDefault();
    if (!inputMessage.trim() || isRunning) return;
    
    const message = inputMessage.trim();
    setInputMessage('');
    executeWorkflow(message);
  };

  // --- Rendering ---

  return (
    <div className="flex h-screen bg-slate-950 text-slate-200 font-sans overflow-hidden select-none">
      
      {/* Sidebar */}
      <aside className="w-64 bg-slate-900 border-r border-slate-800 flex flex-col z-20 shadow-2xl flex-shrink-0">
        <div className="p-4 border-b border-slate-800 bg-slate-900/50 backdrop-blur">
          <div className="flex items-center gap-2 font-bold text-xl bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-purple-400">
            <Bot className="text-blue-400" /> AgentFlow
          </div>
          <p className="text-xs text-slate-500 mt-1">Local Orchestration Studio</p>
        </div>
        
        <div className="p-4 flex-1 overflow-y-auto">
          <div className="text-xs font-bold text-slate-500 uppercase mb-3">Library</div>
          <div className="space-y-2">
            <SidebarItem 
              icon={<Bot size={16}/>} label="AI Agent" color="blue" 
              onClick={() => setNodes(p => [...p, { id: `agent-${Date.now()}`, type: 'agent', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { systemPrompt: 'You are a helpful assistant.', isLocal: true, baseUrl: 'http://172.16.0.2:1234/v1', modelName: 'llama3' } }])} 
            />
            <SidebarItem 
              icon={<Wrench size={16}/>} label="Tool / Skill" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'New Tool', description: 'Capability description...' } }])} 
            />
            <SidebarItem 
              icon={<FileSpreadsheet size={16}/>} label="Google Sheets" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'Google Sheets', description: 'Read data from Google Sheets', toolType: 'google-sheets', sheetId: '', range: 'A1:Z1000', apiKey: '' } }])} 
            />
            <SidebarItem 
              icon={<Wrench size={16}/>} label="Web Search" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'Web Search', description: 'Search the web (Bing). Configure provider and API key for programmatic search', toolType: 'web-search', provider: 'bing', apiKey: '' } }])} 
            />
            <SidebarItem 
              icon={<FileText size={16}/>} label="Local File" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'Local File', description: 'Upload and read Excel / Word / PDF / CSV / TXT', toolType: 'local-file', fileType: 'auto', fileName: '', fileContent: '', loading: false } }])} 
            />
            <SidebarItem 
              icon={<Calculator size={16}/>} label="Calculator" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'Calculator', description: 'Evaluate math expressions', toolType: 'calculator' } }])} 
            />
            <SidebarItem 
              icon={<Play size={16}/>} label="Trigger" color="green" 
              onClick={() => setNodes(p => [...p, { id: `trigger-${Date.now()}`, type: 'trigger', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'Event' } }])} 
            />
          </div>

          <div className="mt-8 p-3 bg-slate-800/50 rounded-lg border border-slate-800">
            <h4 className="text-xs font-bold text-slate-300 mb-2">Controls</h4>
            <div className="text-[10px] text-slate-400 space-y-1">
              <div className="flex justify-between"><span>Pan:</span> <span>Click + Drag BG</span></div>
              <div className="flex justify-between"><span>Move Node:</span> <span>Drag Node</span></div>
              <div className="flex justify-between"><span>Connect:</span> <span>Drag Handles</span></div>
              <div className="flex justify-between"><span>Delete:</span> <span>Click X</span></div>
            </div>
          </div>
        </div>

        {/* Console */}
        <div className="h-64 bg-slate-950 border-t border-slate-800 flex flex-col flex-shrink-0">
          <div className="p-2 border-b border-slate-800 flex justify-between items-center bg-slate-900">
            <span className="text-xs font-bold flex items-center gap-2">
              <Terminal size={12} /> Console
            </span>
            <button onClick={() => setLogs([])} className="text-slate-500 hover:text-red-400"><Trash2 size={12}/></button>
          </div>
          <div className="flex-1 overflow-y-auto p-2 font-mono text-[10px] space-y-1 bg-black/20">
            {logs.length === 0 && <span className="text-slate-700 italic">Waiting for execution...</span>}
            {logs.map((log, i) => (
              <div key={i} className={`flex gap-2 ${getLogColor(log.type)}`}>
                <span className="opacity-30">[{log.timestamp}]</span>
                <span>{log.message}</span>
              </div>
            ))}
            {isRunning && <div className="animate-pulse text-slate-500">_</div>}
          </div>
        </div>
      </aside>

      {/* Main Canvas */}
      <main 
        id="canvas-bg"
        className="flex-1 relative overflow-hidden bg-[#111] cursor-crosshair flex flex-col"
        onMouseDown={handleMouseDown}
        onMouseMove={handleMouseMove}
        onMouseUp={handleMouseUp}
      >
        {/* Grid Pattern */}
        <div 
          className="absolute inset-0 pointer-events-none opacity-20"
          style={{
            backgroundImage: 'radial-gradient(#333 1px, transparent 1px)',
            backgroundSize: `${20 * zoom}px ${20 * zoom}px`,
            backgroundPosition: `${offset.x}px ${offset.y}px`
          }}
        />

        {/* Toolbar */}
        <div className="absolute top-4 right-4 z-20 flex gap-2">
          <button 
            onClick={() => setShowChat(!showChat)}
            className="flex items-center gap-2 px-4 py-2 rounded-lg font-bold shadow-lg transition-all bg-blue-600 hover:bg-blue-500 text-white hover:scale-105"
          >
            <MessageSquare size={16} /> {showChat ? 'Hide Chat' : 'Show Chat'}
          </button>
        </div>

        {/* Graph Container */}
        <div 
          className="absolute inset-0 origin-top-left pointer-events-none"
          style={{ transform: `translate(${offset.x}px, ${offset.y}px) scale(${zoom})` }}
        >
          {/* Edges Layer */}
          <svg className="absolute inset-0 overflow-visible w-full h-full">
            {edges.map(edge => {
              const source = nodes.find(n => n.id === edge.source);
              const target = nodes.find(n => n.id === edge.target);
              if (!source || !target) return null;

              // Simple center-to-center logic for demo, refined for handles roughly
              let start = { x: source.x + (source.type === 'agent' ? 320 : 150), y: source.y + 40 };
              let end = { x: target.x, y: target.y + 40 };

              // Tool Edges are special (vertical or dashed)
              if (edge.type === 'tool') {
                const isTargetTool = target.type === 'tool';
                if (isTargetTool) {
                    start = { x: source.x + 160, y: source.y + 280 }; // Bottom of agent
                    end = { x: target.x + 112, y: target.y }; // Top of tool
                } else {
                    start = { x: source.x + 112, y: source.y }; // Top of tool
                    end = { x: target.x + 160, y: target.y + 280 }; // Bottom of agent
                }
              }

              return (
                <path 
                  key={edge.id}
                  d={generateBezierPath(start, end)}
                  stroke={edge.type === 'tool' ? '#f97316' : '#22c55e'}
                  strokeWidth="3"
                  fill="none"
                  strokeDasharray={edge.type === 'tool' ? "8,4" : "none"}
                  className="opacity-60"
                />
              );
            })}
            
            {/* Active Connection Line */}
            {connecting && (
              <path 
                d={generateBezierPath(
                  { x: connecting.startX, y: connecting.startY }, 
                  screenToWorld(mousePos.x, mousePos.y)
                )}
                stroke={connecting.handleType.includes('tool') ? '#f97316' : '#22c55e'}
                strokeWidth="3"
                fill="none"
                strokeDasharray="4"
                className="opacity-80"
              />
            )}
          </svg>

          {/* Nodes Layer */}
          <div className="pointer-events-auto">
            {nodes.map(node => (
                <Node 
                key={node.id}
                {...node}
                isSelected={selectedNode === node.id}
                onMouseDown={handleNodeMouseDown}
                onHandleMouseDown={handleHandleMouseDown}
                onDelete={handleDeleteNode}
                onUpdate={handleNodeUpdate}
                />
            ))}
          </div>
        </div>

        {/* Chat Panel */}
        {showChat && (
          <div className="absolute bottom-0 left-0 right-0 h-96 bg-slate-900 border-t border-slate-800 flex flex-col z-30 shadow-2xl">
            <div className="p-3 border-b border-slate-800 flex justify-between items-center bg-slate-900/95">
              <div className="flex items-center gap-2">
                <MessageSquare size={16} className="text-blue-400" />
                <span className="text-sm font-bold text-slate-200">Chat</span>
                {isRunning && <span className="text-xs text-slate-500 animate-pulse">Processing...</span>}
              </div>
              <button 
                onClick={() => setShowChat(false)}
                className="text-slate-500 hover:text-slate-300"
              >
                <X size={16} />
              </button>
            </div>
            
            {/* Chat Messages */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-slate-950/50">
              {messages.length === 0 && (
                <div className="text-center text-slate-500 text-sm mt-8">
                  <Bot size={32} className="mx-auto mb-2 opacity-50" />
                  <p>Start a conversation with your local AI agent!</p>
                  <p className="text-xs mt-2">Make sure Ollama or LM Studio is running.</p>
                </div>
              )}
              {messages.map((msg, i) => (
                <div
                  key={i}
                  className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}
                >
                  <div
                    className={`max-w-[80%] rounded-lg p-3 ${
                      msg.role === 'user'
                        ? 'bg-blue-600 text-white'
                        : 'bg-slate-800 text-slate-200 border border-slate-700'
                    }`}
                  >
                    <div className="text-sm whitespace-pre-wrap">{msg.content}</div>
                    {/* Render clickable URL buttons when present */}
                    {msg.urls && msg.urls.length > 0 && (
                      <div className="mt-2 flex flex-col gap-2">
                        {msg.urls.map((u, idx) => (
                          <button
                            key={idx}
                            onClick={() => typeof window !== 'undefined' && window.open(u, '_blank')}
                            className="text-left bg-slate-700 hover:bg-slate-600 px-3 py-1 rounded text-xs text-blue-300"
                          >
                            Open: {u.length > 60 ? u.slice(0, 60) + '...' : u}
                          </button>
                        ))}
                      </div>
                    )}
                    <div className="text-xs opacity-60 mt-1">
                      {msg.timestamp.toLocaleTimeString()}
                    </div>
                  </div>
                </div>
              ))}
              {isRunning && (
                <div className="flex justify-start">
                  <div className="bg-slate-800 text-slate-200 rounded-lg p-3 border border-slate-700">
                    <div className="flex items-center gap-2">
                      <div className="animate-spin">⚙</div>
                      <span className="text-sm">Agent is thinking...</span>
                    </div>
                  </div>
                </div>
              )}
              <div ref={chatEndRef} />
            </div>

            {/* Chat Input */}
            <form onSubmit={handleChatSubmit} className="p-3 border-t border-slate-800 bg-slate-900/95">
              <div className="flex gap-2">
                <input
                  type="text"
                  value={inputMessage}
                  onChange={(e) => setInputMessage(e.target.value)}
                  placeholder="Type your message..."
                  disabled={isRunning}
                  className="flex-1 bg-slate-800 text-slate-200 px-4 py-2 rounded-lg border border-slate-700 focus:border-blue-500 focus:outline-none disabled:opacity-50"
                />
                <button
                  type="submit"
                  disabled={!inputMessage.trim() || isRunning}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all flex items-center gap-2"
                >
                  <Send size={16} />
                </button>
              </div>
            </form>
          </div>
        )}
      </main>
    </div>
  );
}

const SidebarItem = ({ icon, label, color, onClick }) => {
  const colors = {
    blue: "bg-blue-500/10 border-blue-500/30 text-blue-400 hover:border-blue-400 hover:bg-blue-500/20",
    orange: "bg-orange-500/10 border-orange-500/30 text-orange-400 hover:border-orange-400 hover:bg-orange-500/20",
    green: "bg-green-500/10 border-green-500/30 text-green-400 hover:border-green-400 hover:bg-green-500/20",
  }
  return (
    <div 
      onClick={onClick}
      className={`p-3 rounded-lg border cursor-pointer flex items-center gap-3 transition-all ${colors[color]}`}
    >
      {icon}
      <span className="text-xs font-medium">{label}</span>
      <Plus size={12} className="ml-auto opacity-50" />
    </div>
  )
}

const getLogColor = (type) => {
  switch (type) {
    case 'error': return 'text-red-400 font-bold';
    case 'warn': return 'text-yellow-400';
    case 'success': return 'text-green-400 font-bold';
    case 'agent': return 'text-blue-400';
    case 'tool': return 'text-orange-400';
    case 'process': return 'text-purple-400';
    default: return 'text-slate-300';
  }
};

