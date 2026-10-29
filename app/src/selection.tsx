import React, { useDeferredValue, useMemo, useState } from 'react';
import { Search } from 'lucide-react';

export type Book = 'quantity' | 'financial';
export type Article = { id: string; book: Book; code: string; name: string; group: string };
export type Catalog = { months: string[]; branches: string[]; articles: Article[] };
export const bookTitle: Record<Book, string> = { quantity: 'Количественный бюджет', financial: 'Финансовый бюджет' };
const groupOrder = ['Основные показатели', 'Лекарственные препараты', 'Изделия медицинского назначения', 'Расходные медицинские материалы', 'Переменные косвенные затраты', 'Постоянные затраты', 'Прочие статьи'];

export function ChoicePanel({ title, values, selected, onChange, columns = 1 }: { title: string; values: string[]; selected: string[]; onChange: (items: string[]) => void; columns?: 1 | 2 }) {
  function toggle(value: string) { onChange(selected.includes(value) ? selected.filter(x => x !== value) : [...selected, value]); }
  return <fieldset className="option-panel"><legend>{title} <span>{selected.length}/{values.length}</span></legend>
    {title === 'Филиалы' && <button className="text-button" onClick={() => onChange(selected.length === values.length ? [] : [...values])}>{selected.length === values.length ? 'Снять выбор' : 'Выбрать все'}</button>}
    <div className={`option-list ${columns === 2 ? 'months' : 'branches'}`}>{values.map(value => <label key={value} className="check-row"><input type="checkbox" checked={selected.includes(value)} onChange={() => toggle(value)}/><span>{columns === 2 ? value[0].toUpperCase() + value.slice(1) : value}</span></label>)}</div>
  </fieldset>;
}

export function ArticleTable({ items, selected, onChange }: { items: Article[]; selected: string[]; onChange: (items: string[]) => void }) {
  const [query, setQuery] = useState('');
  const [source, setSource] = useState<'all' | Book>('all');
  const deferredQuery = useDeferredValue(query);
  const visible = useMemo(() => items.filter(item =>
    (source === 'all' || item.book === source) &&
    `${item.code} ${item.name}`.toLocaleLowerCase('ru').includes(deferredQuery.trim().toLocaleLowerCase('ru'))
  ), [items, source, deferredQuery]);
  const grouped = groupOrder.map(group => [group, visible.filter(item => item.group === group)] as const).filter(([, rows]) => rows.length);
  function toggle(id: string) { onChange(selected.includes(id) ? selected.filter(x => x !== id) : [...selected, id]); }
  function toggleGroup(rows: Article[]) {
    const ids = rows.map(x => x.id);
    onChange(ids.every(id => selected.includes(id)) ? selected.filter(id => !ids.includes(id)) : [...new Set([...selected, ...ids])]);
  }
  return <div className="article-panel"><div className="article-head"><div><h3>Статьи</h3><p>Выбрано {selected.length}. Код и название считаются отдельной позицией для каждой книги.</p></div><div className="article-tools"><label className="search"><Search size={17} aria-hidden="true"/><span className="sr-only">Поиск статьи</span><input name="article-search" autoComplete="off" value={query} onChange={e => setQuery(e.target.value)} placeholder="Код или название…"/></label><select aria-label="Источник статьи" value={source} onChange={e => setSource(e.target.value as typeof source)}><option value="all">Обе книги</option><option value="quantity">Количественная</option><option value="financial">Финансовая</option></select></div></div>
    <div className="table-scroll"><table><thead><tr><th className="check-cell"><span className="sr-only">Выбор</span></th><th>Код</th><th>Наименование статьи</th><th>Книга</th></tr></thead><tbody>{grouped.map(([group, rows]) => <React.Fragment key={group}><tr className="group-row"><td colSpan={4}><span>{group}</span><button className="text-button" onClick={() => toggleGroup(rows)}>{rows.every(x => selected.includes(x.id)) ? 'Снять' : 'Выбрать'} ({rows.length})</button></td></tr>{rows.map(item => <tr key={item.id}><td className="check-cell"><input type="checkbox" aria-label={`Выбрать ${item.code} ${item.name}, ${bookTitle[item.book]}`} checked={selected.includes(item.id)} onChange={() => toggle(item.id)}/></td><td className="code">{item.code}</td><td>{item.name}</td><td><span className="book-badge">{item.book === 'quantity' ? 'Кол.' : 'Фин.'}</span></td></tr>)}</React.Fragment>)}</tbody></table>{visible.length === 0 && <div className="empty">По запросу статьи не найдены.</div>}</div>
  </div>;
}
