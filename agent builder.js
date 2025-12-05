import React, { useState, useRef, useCallback, useEffect } from 'react';
import { 
  Bot, 
  Wrench, 
  Play, 
  Database, 
  Settings, 
  Trash2, 
  Terminal, 
  Save, 
  Plus, 
  X,
  Search,
  MousePointer2,
  Server,
  Cloud
} from 'lucide-react';

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
                    <label className="text-[9px] uppercase text-slate-500 font-bold mb-1 block">Local Endpoint (Ollama/LM Studio)</label>
                    <input 
                      className="w-full bg-slate-950 text-xs text-slate-300 p-2 rounded border border-slate-800 focus:border-blue-500 outline-none font-mono"
                      defaultValue={data.baseUrl || "http://localhost:11434/v1"}
                      placeholder= "http://localhost:11434/v1"
                      onChange={(e) => handleInputChange('baseUrl', e.target.value)}
                      onMouseDown={(e) => e.stopPropagation()}
                    />
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
              <Wrench size={14} className="text-orange-400" />
              <span className="text-xs font-bold text-orange-100">{data.label}</span>
            </div>
            <button onClick={(e) => { e.stopPropagation(); onDelete(id); }} className="text-orange-400/50 hover:text-orange-400"><X size={12}/></button>
          </div>
          <div className="p-2 bg-slate-900 text-[10px] text-slate-400">
            {data.description}
          </div>
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
  { id: 'agent-1', type: 'agent', x: 450, y: 150, data: { systemPrompt: 'You are a research assistant.', isLocal: true, baseUrl: 'http://localhost:11434/v1', modelName: 'llama3' } },
  { id: 'tool-1', type: 'tool', x: 450, y: 600, data: { label: 'Google Search', description: 'Search the web for info.' } },
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
        if (connecting.handleType === 'tools' || connecting.handleType === 'tool-out') {
          type = 'tool';
        }

        const newEdge = {
          id: `e-${Date.now()}`,
          source: connecting.nodeId,
          target: targetNode.id,
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

  // --- Simulation Engine ---
  const runSimulation = async () => {
    if (isRunning) return;
    setIsRunning(true);
    setLogs([]);
    addLog("system", "Initializing Local-First Engine...");

    const trigger = nodes.find(n => n.type === 'trigger');
    if (!trigger) {
      addLog("error", "Error: No Trigger node found.");
      setIsRunning(false);
      return;
    }
    
    addLog("info", `Trigger Event: [${trigger.data.label}] received input.`);
    await wait(500);

    const flowEdge = edges.find(e => e.source === trigger.id && e.type === 'flow');
    if (!flowEdge) {
      addLog("warn", "Warning: Trigger not connected to any Agent.");
      setIsRunning(false);
      return;
    }

    const agent = nodes.find(n => n.id === flowEdge.target);
    addLog("process", `Routing to Agent: [${agent.id}]`);
    await wait(500);

    // Local vs Cloud Logic check
    if (agent.data.isLocal) {
        addLog("system", `Connecting to Local API: ${agent.data.baseUrl}`);
        addLog("system", `Loading Model: ${agent.data.modelName}...`);
        
        // Simulate network check
        try {
            addLog("system", `Checking connection to ${agent.data.baseUrl}...`);
            await wait(800);
            addLog("success", "Local Server (Ollama/LM Studio) connected.");
        } catch (e) {
            addLog("error", "Failed to connect to Local Server.");
        }
    } else {
        addLog("system", `Connecting to Cloud API (OpenAI/Anthropic)...`);
    }

    addLog("agent", `Agent Context: "${agent.data.systemPrompt}"`);
    await wait(800);

    const connectedTools = edges
      .filter(e => (e.target === agent.id || e.source === agent.id) && e.type === 'tool')
      .map(e => nodes.find(n => n.id === (e.source === agent.id ? e.target : e.source)))
      .filter(n => n && n.type === 'tool');

    if (connectedTools.length > 0) {
      addLog("info", `Agent identified ${connectedTools.length} connected tools.`);
      
      for (const tool of connectedTools) {
        await wait(600);
        addLog("agent", `Thinking (Chain-of-Thought): "User asked for info, using ${tool.data.label}..."`);
        await wait(800);
        addLog("tool", `POST /v1/chat/completions -> Tool Call: [${tool.data.label}]`);
        await wait(1000);
        addLog("process", `Tool returned results. Feeding back to ${agent.data.modelName || "model"}.`);
      }
    }

    await wait(800);
    addLog("success", "Final Response generated.");
    setIsRunning(false);
  };

  const addLog = (type, message) => {
    setLogs(prev => [...prev, { type, message, timestamp: new Date().toLocaleTimeString() }]);
  };
  
  const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  // --- Rendering ---

  return (
    <div className="flex h-screen bg-slate-950 text-slate-200 font-sans overflow-hidden select-none">
      
      {/* Sidebar */}
      <aside className="w-64 bg-slate-900 border-r border-slate-800 flex flex-col z-20 shadow-2xl">
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
              onClick={() => setNodes(p => [...p, { id: `agent-${Date.now()}`, type: 'agent', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { systemPrompt: 'You are a helpful assistant.', isLocal: true, baseUrl: 'http://localhost:11434/v1', modelName: 'llama3' } }])} 
            />
            <SidebarItem 
              icon={<Wrench size={16}/>} label="Tool / Skill" color="orange" 
              onClick={() => setNodes(p => [...p, { id: `tool-${Date.now()}`, type: 'tool', x: -offset.x/zoom + 100, y: -offset.y/zoom + 100, data: { label: 'New Tool', description: 'Capability description...' } }])} 
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
        className="flex-1 relative overflow-hidden bg-[#111] cursor-crosshair"
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
            onClick={runSimulation}
            disabled={isRunning}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold shadow-lg transition-all ${isRunning ? 'bg-slate-800 text-slate-500 cursor-not-allowed' : 'bg-green-600 hover:bg-green-500 text-white hover:scale-105'}`}
          >
            {isRunning ? 'Running...' : <><Play size={16} fill="currentColor" /> Run Flow</>}
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