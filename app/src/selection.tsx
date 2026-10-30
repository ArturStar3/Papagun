import { useDeferredValue, useMemo, useRef, useState, type MouseEvent } from 'react';
import { Search, Trash2 } from 'lucide-react';

export type Book = 'quantity' | 'financial';
export type Article = { id: string; code: string; name: string; group: string };
export type Catalog = { months: string[]; availableMonths: string[]; branches: string[]; availableBranches: string[]; articles: Article[] };
export type ArticleGroup = { id: string; title: string; articles: string[] };
export const bookTitle: Record<Book, string> = { quantity: 'Количественный бюджет', financial: 'Финансовый бюджет' };
const categories = ['Лекарственные препараты', 'Изделия медицинского назначения', 'Расходные медицинские материалы'];

export function ChoicePanel({ title, values, available = values, selected, onChange }: {
  title: string; values: string[]; available?: string[]; selected: string[]; onChange: (items: string[]) => void;
}) {
  const anchor = useRef<number | null>(null);
  const enabled = values.filter(value => available.includes(value));
  function choose(index: number, event: MouseEvent<HTMLButtonElement>) {
    const value = values[index];
    if (!available.includes(value)) return;
    if (event.shiftKey && anchor.current !== null) {
      const range = values.slice(Math.min(anchor.current, index), Math.max(anchor.current, index) + 1).filter(x => available.includes(x));
      onChange([...new Set([...selected, ...range])]);
    } else if (event.ctrlKey || event.metaKey) {
      onChange(selected.includes(value) ? selected.filter(x => x !== value) : [...selected, value]);
      anchor.current = index;
    } else {
      onChange([value]);
      anchor.current = index;
    }
  }
  return <fieldset className="tile-fieldset"><legend>{title} <span>{selected.length}/{enabled.length}</span></legend>
    <div className="tile-actions"><button type="button" className="text-button" onClick={() => onChange(enabled)}>Выбрать все</button><button type="button" className="text-button" onClick={() => onChange([])}>Снять выбор</button></div>
    <div className={`tile-grid ${title === 'Месяцы' ? 'month-tiles' : 'branch-tiles'}`}>
      {values.map((value, index) => <button type="button" key={value} className={`choice-tile ${selected.includes(value) ? 'is-selected' : ''}`}
        disabled={!available.includes(value)} aria-pressed={selected.includes(value)} title={!available.includes(value) ? 'Нет полного набора данных для отчёта' : value}
        onClick={event => choose(index, event)}>{value[0].toUpperCase() + value.slice(1)}</button>)}
    </div>
    <p className="tile-hint">Щелчок — один пункт · Shift + щелчок — диапазон · Ctrl + щелчок — добавить или снять</p>
  </fieldset>;
}

export function ArticleTable({ items, selected, groups, onChange, onGroupsChange }: {
  items: Article[]; selected: string[]; groups: ArticleGroup[];
  onChange: (items: string[]) => void; onGroupsChange: (items: ArticleGroup[]) => void;
}) {
  const [category, setCategory] = useState(categories[0]);
  const [query, setQuery] = useState('');
  const [groupTitle, setGroupTitle] = useState('');
  const deferredQuery = useDeferredValue(query);
  const groupIds = new Set(groups.flatMap(group => group.articles));
  const current = useMemo(() => items.filter(item => item.group === category &&
    `${item.code} ${item.name}`.toLocaleLowerCase('ru').includes(deferredQuery.trim().toLocaleLowerCase('ru'))), [items, category, deferredQuery]);
  const selectedHere = selected.filter(id => items.some(item => item.id === id && item.group === category));
  const selectableHere = current.filter(item => !groupIds.has(item.id) && !['211', '212', '213'].includes(item.code)).map(item => item.id);
  const selectedVisible = selectableHere.filter(id => selected.includes(id));
  function selectAllHere() { onChange([...new Set([...selected, ...selectableHere])]); }
  function clearHere() {
    const visibleIds = new Set(selectableHere);
    onChange(selected.filter(id => !visibleIds.has(id)));
  }
  function toggle(id: string) { onChange(selected.includes(id) ? selected.filter(x => x !== id) : [...selected, id]); }
  function createGroup() {
    const title = groupTitle.trim();
    if (!title || selectedHere.length < 2) return;
    onGroupsChange([...groups, { id: crypto.randomUUID(), title, articles: selectedHere }]);
    onChange(selected.filter(id => !selectedHere.includes(id)));
    setGroupTitle('');
  }
  return <div className="article-panel"><div className="article-head"><div><h3>Статьи и показатели</h3><p>Статистика и расходы добавляются автоматически. Здесь выберите статьи для количественной, финансовой и квартальной страниц.</p></div></div>
    <div className="category-tabs" role="group" aria-label="Раздел статей">{categories.map(name => <button key={name} aria-pressed={category === name}
      className={category === name ? 'active' : ''} onClick={() => { setCategory(name); setQuery(''); }}>{name}<span>{items.filter(item => item.group === name).length}</span></button>)}</div>
    <div className="category-content"><div className="category-tools"><label className="search"><Search size={17} aria-hidden="true"/><span className="sr-only">Поиск статьи</span><input name="article-search" autoComplete="off" value={query} onChange={event => setQuery(event.target.value)} placeholder="Код или название…"/></label>
      <div className="category-selection"><span className="category-count">Выбрано в разделе: {selectedHere.length}</span>
        <div className="category-actions"><button type="button" className="text-button" onClick={selectAllHere} disabled={!selectableHere.length || selectedVisible.length === selectableHere.length}>{deferredQuery.trim() ? 'Выбрать найденные' : 'Выбрать все'}</button>
          <button type="button" className="text-button" onClick={clearHere} disabled={!selectedVisible.length}>Снять выбор</button></div></div></div>
      <div className="table-scroll"><table><thead><tr><th className="check-cell"><span className="sr-only">Выбор</span></th><th>Код</th><th>Наименование статьи</th></tr></thead><tbody>{current.map(item => <tr key={item.id}><td className="check-cell"><input type="checkbox" aria-label={`Выбрать ${item.code} ${item.name}`}
        checked={selected.includes(item.id) || groupIds.has(item.id)} disabled={groupIds.has(item.id) || ['211', '212', '213'].includes(item.code)} onChange={() => toggle(item.id)}/></td><td className="code">{item.code}</td><td>{item.name}{groupIds.has(item.id) && <span className="in-group">В группе</span>}</td></tr>)}</tbody></table>{current.length === 0 && <div className="empty">В этом разделе статьи не найдены.</div>}</div>
      <div className="group-creator"><div><strong>Объединить в один показатель</strong><p>Выберите не менее двух статей этого раздела. Значения сложатся по филиалу и периоду.</p></div><label><span className="sr-only">Заголовок группы на слайде</span><input name="group-title" autoComplete="off" value={groupTitle} onChange={event => setGroupTitle(event.target.value)} placeholder="Заголовок на слайде…" maxLength={100}/></label><button className="secondary" type="button" onClick={createGroup} disabled={selectedHere.length < 2 || !groupTitle.trim()}>Создать группу</button></div>
    </div>
    {groups.length > 0 && <div className="created-groups"><h4>Составные показатели</h4>{groups.map(group => <div className="created-group" key={group.id}><div><strong>{group.title}</strong><span>{group.articles.map(id => items.find(item => item.id === id)?.code).join(' + ')}</span></div><button type="button" className="icon-button" title="Удалить группу" aria-label={`Удалить группу ${group.title}`} onClick={() => onGroupsChange(groups.filter(item => item.id !== group.id))}><Trash2 size={17}/></button></div>)}</div>}
  </div>;
}
