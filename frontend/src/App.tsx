import { useState, useEffect, useRef } from 'react';

const API_BASE = 'http://localhost:5000/api';

type AutonomyLevel = "Autônomo" | "Parcialmente autônomo" | "Apoiado" | "Não satisfatório";

interface PendingStudent {
  nome: string;
}

export default function App() {
  const [logs, setLogs] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [pendingStudents, setPendingStudents] = useState<PendingStudent[]>([]);
  const [autonomyAnswers, setAutonomyAnswers] = useState<Record<string, AutonomyLevel>>({});
  const [isWaitingClassroom, setIsWaitingClassroom] = useState(false);
  const [classroomUrl, setClassroomUrl] = useState<string>('');
  const logsEndRef = useRef<HTMLDivElement>(null);

  const fetchStatus = async () => {
    try {
      const res = await fetch(`${API_BASE}/status`);
      const data = await res.json();
      setIsProcessing(data.processando);
    } catch (e) {
      // Ignore
    }
  };

  const fetchLogs = async () => {
    try {
      const res = await fetch(`${API_BASE}/logs`);
      const data = await res.json();
      if (data.logs) {
        // Corrige a formatação de \n vinda do backend Flask
        const cleanLogs = data.logs.replace(/\\n/g, '\n');
        setLogs((prev) => prev + cleanLogs);
      }
    } catch (e) {
      // Ignore
    }
  };

  useEffect(() => {
    const interval = setInterval(() => {
      fetchStatus();
      fetchLogs();
    }, 1000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    if (logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs]);

  const handleExtrair = async () => {
    await fetch(`${API_BASE}/extrair`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url: classroomUrl })
    });
    setLogs((prev) => prev + '\n[SISTEMA] Iniciando extração, abrindo o navegador...\n');
    setIsProcessing(true);
    if (!classroomUrl) {
      setIsWaitingClassroom(true);
    } else {
      setIsWaitingClassroom(false);
    }
  };

  const handleConfirmExtract = async () => {
    await fetch(`${API_BASE}/extrair/confirmar`, { method: 'POST' });
    setIsWaitingClassroom(false);
  };

  const handlePreencher = async () => {
    const res = await fetch(`${API_BASE}/preencher/check`);
    const data = await res.json();
    if (data.pendentes && data.pendentes.length > 0) {
      setPendingStudents(data.pendentes);
    } else {
      await executePreenchimento({});
    }
  };

  const executePreenchimento = async (answers: Record<string, string>) => {
    await fetch(`${API_BASE}/preencher/execute`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ autonomias: answers })
    });
    setLogs((prev) => prev + '\\n[SISTEMA] Iniciando preenchimento...\\n');
    setIsProcessing(true);
    setPendingStudents([]);
  };

  const handleStop = async () => {
    await fetch(`${API_BASE}/stop`, { method: 'POST' });
  };

  const handleAutonomyChange = (student: string, level: AutonomyLevel) => {
    setAutonomyAnswers(prev => ({ ...prev, [student]: level }));
  };

  const submitAutonomy = () => {
    // Todos os alunos listados têm graus válidos e precisam de autonomia definida
    const allAnswered = pendingStudents.every(s => autonomyAnswers[s.nome]);
    if (!allAnswered) {
      alert("Por favor, preencha a autonomia de todos os alunos listados.");
      return;
    }
    executePreenchimento(autonomyAnswers);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-8 font-sans">
      <div className="max-w-5xl mx-auto space-y-6">
        <header className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-slate-800">Automação de Rubricas</h1>
            <p className="text-slate-500 mt-1">Extraia do Classroom e preencha no Excel com facilidade.</p>
          </div>
          <div className="flex items-center space-x-3">
            <span className="relative flex h-3 w-3">
              {isProcessing && <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>}
              <span className={`relative inline-flex rounded-full h-3 w-3 ${isProcessing ? 'bg-blue-500' : 'bg-slate-300'}`}></span>
            </span>
            <span className="text-sm font-medium text-slate-600">{isProcessing ? 'Processando...' : 'Pronto'}</span>
          </div>
        </header>

        {/* Seção 1: Configuração da URL e Ações (Stack Horizontal) */}
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 space-y-4">
          <div className="space-y-2">
            <label className="text-sm font-semibold text-slate-700 block">URL da Atividade (Classroom)</label>
            <input
              type="text"
              placeholder="Cole a URL da atividade aqui..."
              value={classroomUrl}
              onChange={(e) => setClassroomUrl(e.target.value)}
              className="w-full text-sm p-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-slate-800"
            />
            <span className="text-xs text-slate-400 block">Se informado, roda em 2º plano (oculto). Deixe vazio para manual/visível.</span>
          </div>

          <div className="flex flex-wrap gap-3">
            {!isWaitingClassroom ? (
              <button
                onClick={handleExtrair}
                disabled={isProcessing}
                className="flex-1 min-w-[150px] bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-semibold py-3 px-6 rounded-xl transition-all shadow-sm active:scale-[0.98]"
              >
                1. Extrair Rubricas
              </button>
            ) : (
              <button
                onClick={handleConfirmExtract}
                className="flex-1 min-w-[200px] bg-amber-500 hover:bg-amber-600 text-white font-semibold py-3 px-6 rounded-xl transition-all shadow-sm active:scale-[0.98] animate-pulse"
              >
                Confirmar Página do Classroom
              </button>
            )}
            <button
              onClick={handlePreencher}
              disabled={isProcessing}
              className="flex-1 min-w-[150px] bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white font-semibold py-3 px-6 rounded-xl transition-all shadow-sm active:scale-[0.98]"
            >
              2. Preencher Planilha
            </button>
            <button
              onClick={handleStop}
              disabled={!isProcessing}
              className="flex-1 min-w-[100px] bg-rose-100 hover:bg-rose-200 text-rose-700 disabled:bg-slate-100 disabled:text-slate-400 disabled:cursor-not-allowed font-semibold py-3 px-6 rounded-xl transition-all"
            >
              Parar
            </button>
          </div>
        </div>

        {/* Seção 2: Console Output em largura total */}
        <div className="bg-slate-900 rounded-2xl shadow-xl overflow-hidden flex flex-col h-[550px] w-full">
          <div className="bg-slate-800 px-4 py-3 border-b border-slate-700 flex items-center">
            <div className="flex space-x-2">
              <div className="w-3 h-3 rounded-full bg-rose-500"></div>
              <div className="w-3 h-3 rounded-full bg-amber-500"></div>
              <div className="w-3 h-3 rounded-full bg-emerald-500"></div>
            </div>
            <span className="ml-4 text-xs font-mono text-slate-400">console output</span>
          </div>
          <div className="p-4 flex-1 overflow-y-auto font-mono text-sm text-emerald-400 whitespace-pre-wrap">
            {logs || 'Aguardando inicialização...\n'}
            <div ref={logsEndRef} />
          </div>
        </div>
      </div>

      {pendingStudents.length > 0 && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-slate-800">Definir Autonomia</h2>
              <p className="text-slate-500 mt-1">Os alunos abaixo têm critérios avaliados. Defina a autonomia de cada um.</p>
            </div>
            
            <div className="p-6 overflow-y-auto flex-1 space-y-6">
              {pendingStudents.map(student => (
                <div key={student.nome} className="space-y-3 bg-slate-50 p-4 rounded-xl border border-slate-100">
                  <h3 className="font-semibold text-slate-800">{student.nome}</h3>
                  <div className="grid grid-cols-2 gap-2">
                    {["Autônomo", "Parcialmente autônomo", "Apoiado", "Não satisfatório"].map(level => (
                      <label key={level} className={`
                        flex items-center p-3 rounded-lg border cursor-pointer transition-all
                        ${autonomyAnswers[student.nome] === level ? 'bg-blue-50 border-blue-500 text-blue-700 ring-1 ring-blue-500' : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300'}
                      `}>
                        <input
                          type="radio"
                          name={student.nome}
                          value={level}
                          checked={autonomyAnswers[student.nome] === level}
                          onChange={() => handleAutonomyChange(student.nome, level as AutonomyLevel)}
                          className="sr-only"
                        />
                        <span className="text-sm font-medium">{level}</span>
                      </label>
                    ))}
                  </div>
                </div>
              ))}
            </div>

            <div className="p-6 border-t border-slate-100 bg-slate-50 flex justify-end space-x-3">
              <button 
                onClick={() => setPendingStudents([])}
                className="px-6 py-2.5 rounded-xl font-medium text-slate-600 hover:bg-slate-200 transition-colors"
              >
                Cancelar
              </button>
              <button 
                onClick={submitAutonomy}
                className="px-6 py-2.5 rounded-xl font-medium bg-blue-600 hover:bg-blue-700 text-white shadow-sm transition-colors"
              >
                Confirmar e Preencher
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
