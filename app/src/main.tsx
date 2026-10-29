import { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { invoke } from '@tauri-apps/api/core';
import { open, save } from '@tauri-apps/plugin-dialog';
import { Check, ChevronRight, FileSpreadsheet, FileText, FolderOpen, Layers3, Moon, Sun, X } from 'lucide-react';
import { ArticleTable, ChoicePanel, bookTitle, type ArticleGroup, type Book, type Catalog } from './selection';
import './style.css';

type Result = { output: string; slides: number; articles: number };

function App() {
  const [theme, setTheme] = useState<'light' | 'dark'>(() => localStorage.getItem('papagun-theme') === 'dark' ? 'dark' : 'light');
  const [files, setFiles] = useState<Record<Book, string>>({ quantity: '', financial: '' });
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [months, setMonths] = useState<string[]>([]);
  const [branches, setBranches] = useState<string[]>([]);
  const [articles, setArticles] = useState<string[]>([]);
  const [groups, setGroups] = useState<ArticleGroup[]>([]);
  const [busy, setBusy] = useState<'inspect' | 'generate' | null>(null);
  const [error, setError] = useState('');
  const [result, setResult] = useState<Result | null>(null);
  const [confirm, setConfirm] = useState(false);
  const sourcesRef = useRef<HTMLElement>(null);
  const filtersRef = useRef<HTMLElement>(null);
  const exportRef = useRef<HTMLElement>(null);
  const generateButtonRef = useRef<HTMLButtonElement>(null);

  function closeConfirm() {
    setConfirm(false);
    requestAnimationFrame(() => generateButtonRef.current?.focus());
  }

  function toggleTheme() {
    const next = theme === 'light' ? 'dark' : 'light';
    setTheme(next); localStorage.setItem('papagun-theme', next);
  }
  async function pick(book: Book) {
    const selected = await open({ multiple: false, filters: [{ name: 'Excel', extensions: ['xlsx'] }] });
    if (typeof selected === 'string') {
      setFiles(prev => ({ ...prev, [book]: selected }));
      setCatalog(null); setMonths([]); setBranches([]); setArticles([]); setGroups([]); setResult(null); setError('');
    }
  }
  async function inspect() {
    setBusy('inspect'); setError(''); setResult(null);
    try {
      const next = await invoke<Catalog>('engine', { action: 'inspect', payload: files });
      setCatalog(next); setMonths(next.availableMonths); setBranches(next.availableBranches); setArticles([]); setGroups([]);
      requestAnimationFrame(() => filtersRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }));
    } catch (e) { setError(String(e)); }
    finally { setBusy(null); }
  }
  async function generate() {
    closeConfirm();
    const output = await save({ defaultPath: 'Papagun_отчёт.pptx', filters: [{ name: 'PowerPoint', extensions: ['pptx'] }] });
    if (!output) return;
    setBusy('generate'); setError(''); setResult(null);
    try {
      const next = await invoke<Result>('engine', { action: 'generate', payload: { ...files, months, branches, articles, groups, output } });
      setResult(next);
      requestAnimationFrame(() => exportRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }));
    } catch (e) { setError(String(e)); }
    finally { setBusy(null); }
  }
  const canGenerate = Boolean(catalog && months.length && branches.length && !busy);

  return <div className={`app ${theme}`}>
    <aside className="sidebar" aria-label="Навигация">
      <div className="brand"><div className="brand-mark"><Layers3 size={21}/></div><div><strong>Papagun</strong><span>Конструктор отчётов</span></div></div>
      <nav className="nav">
        <button onClick={() => sourcesRef.current?.scrollIntoView({ behavior: 'smooth' })}><FileSpreadsheet size={18}/> Источники <ChevronRight size={15}/></button>
        <button onClick={() => filtersRef.current?.scrollIntoView({ behavior: 'smooth' })} disabled={!catalog}><Layers3 size={18}/> Выбор данных <ChevronRight size={15}/></button>
        <button onClick={() => exportRef.current?.scrollIntoView({ behavior: 'smooth' })} disabled={!catalog}><FileText size={18}/> Презентация <ChevronRight size={15}/></button>
      </nav>
      <div className="sidebar-bottom"><div className="offline-dot"/> Локальная обработка</div>
    </aside>
    <main className="workspace">
      <header className="topbar"><span>Рабочая область</span><button className="icon-button" onClick={toggleTheme} aria-label={theme === 'light' ? 'Включить тёмную тему' : 'Включить светлую тему'} title="Сменить тему">{theme === 'light' ? <Moon size={19}/> : <Sun size={19}/>}</button></header>
      <div className="content">
        <div className="heading"><div><h1>Подготовка презентации</h1><p>Загрузите две книги бюджета, выберите данные и сохраните отчёт PowerPoint.</p></div><span className="version">NL / 2026</span></div>
        <span className="sr-only" role="status">{busy === 'inspect' ? 'Чтение книг' : busy === 'generate' ? 'Создание презентации' : ''}</span>
        <section className="section" ref={sourcesRef} aria-labelledby="sources-title">
          <div className="section-title"><span className="step">1</span><div><h2 id="sources-title">Исходные книги</h2><p>Выберите количественный и финансовый бюджет в формате .xlsx.</p></div></div>
          <div className="source-grid">{(['quantity', 'financial'] as Book[]).map(book => <div className="source-card" key={book}>
            <div className="source-icon"><FileSpreadsheet size={22}/></div><div className="source-copy"><h3>{bookTitle[book]}</h3><p title={files[book]}>{files[book] ? files[book].split(/[\\/]/).pop() : 'Файл не выбран'}</p></div>
            <button className="secondary" onClick={() => pick(book)} disabled={!!busy}><FolderOpen size={16}/> Выбрать</button>
          </div>)}</div>
          <div className="section-actions"><button className="primary" onClick={inspect} disabled={!files.quantity || !files.financial || !!busy}>{busy === 'inspect' ? 'Чтение книг…' : catalog ? 'Обновить данные' : 'Загрузить данные'}</button>{catalog && <span className="inline-success"><Check size={16}/> Найдено {catalog.articles.length} позиций</span>}</div>
        </section>
        {catalog && <section className="section" ref={filtersRef} aria-labelledby="selection-title">
          <div className="section-title"><span className="step">2</span><div><h2 id="selection-title">Данные отчёта</h2><p>Выберите месяцы и филиалы. Первые три слайда с аппаратами, пациентами и процедурами строятся автоматически; статьи ниже добавляют тематические слайды.</p></div></div>
          <div className="selection-grid">
            <ChoicePanel title="Месяцы" values={catalog.months} available={catalog.availableMonths} selected={months} onChange={items => { setMonths(items); setResult(null); }}/>
            <ChoicePanel title="Филиалы" values={catalog.branches} available={catalog.availableBranches} selected={branches} onChange={items => { setBranches(items); setResult(null); }}/>
          </div>
          <ArticleTable items={catalog.articles} selected={articles} groups={groups} onChange={items => { setArticles(items); setResult(null); }} onGroupsChange={items => { setGroups(items); setResult(null); }}/></section>}
        {catalog && <section className="section result-section" ref={exportRef} aria-labelledby="export-title"><div className="section-title"><span className="step">3</span><div><h2 id="export-title">Презентация</h2><p>Порядок: три статистических слайда, выбранные разделы статей, затем два слайда расходов. Диаграммы остаются редактируемыми.</p></div></div><div className="export-bar"><div><strong>{articles.length + groups.length} показателей из статей</strong><span>{months.length} мес. · {branches.length} фил. · 5 обязательных слайдов</span></div><button className="primary" ref={generateButtonRef} onClick={() => setConfirm(true)} disabled={!canGenerate}>{busy === 'generate' ? 'Создание презентации…' : 'Создать презентацию'}</button></div>{result && <div className="success" role="status"><Check size={20}/><div><strong>Презентация сохранена</strong><span>{result.slides} слайдов · {result.output}</span></div></div>}</section>}
        {error && <div className="error" role="alert"><span>{error}</span><button aria-label="Закрыть сообщение" onClick={() => setError('')}><X size={17}/></button></div>}
      </div>
    </main>
    {confirm && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) closeConfirm(); }}><div className="modal" role="dialog" aria-modal="true" aria-labelledby="confirm-title" onKeyDown={e => {
      if (e.key === 'Escape') closeConfirm();
      if (e.key === 'Tab') {
        const buttons = Array.from(e.currentTarget.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'));
        const first = buttons[0], last = buttons[buttons.length - 1];
        if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    }}><button className="modal-close" aria-label="Закрыть" onClick={() => closeConfirm()}><X size={18}/></button><div className="modal-icon"><FileText size={22}/></div><h2 id="confirm-title">Создать презентацию?</h2><p>Будут обработаны {articles.length + groups.length} показателей из статей, статистика и расходы для {branches.length} филиалов и {months.length} месяцев. Следующим шагом выберите место сохранения.</p><div className="modal-actions"><button className="secondary" onClick={() => closeConfirm()}>Отмена</button><button className="primary" autoFocus onClick={generate}>Выбрать файл</button></div></div></div>}
  </div>;
}

createRoot(document.getElementById('root')!).render(<App/>);
