import { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  Activity, Archive, ArrowUpRight, BarChart3, BriefcaseBusiness, Building2,
  ChevronDown, CircleHelp, FileDown, FileUp, History, LayoutDashboard,
  GitBranch, Menu, Plus, Search, Settings2, ShieldCheck, UserRound, UsersRound, X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import './styles.css';

type Department = { id: string; name: string; active: boolean; areas: Area[] };
type Area = { id: string; name: string; active: boolean; positions: Position[] };
type Position = { id: string; name: string; active: boolean; personId?: string; assignmentConflict?: boolean };
type Person = { id: string; fullName: string; identification?: string; active: boolean; assignmentConflict?: boolean };
type View = 'dashboard' | 'organization' | 'tree' | 'people' | 'assignments' | 'vacancies' | 'unassigned' | 'history' | 'settings';
type ImportedRow = { department: string; area: string; position: string; person: string };

const baseDepartments = [
  'Misiones Locales', 'Misiones Extranjeras', 'Educación Cristiana', 'Decom',
  'Conquistadores Pentecostales', 'Damas Dorcas', 'Caballeros', 'Familia',
  'Obra Social', 'Ornato y Embellecimiento', 'Seguridad y Prevención de Riesgos Eclesial',
  'Construcción', 'Mantenimiento y Cuidados Locativos', 'Intercesión',
  'Atención y Protocolo', 'Junta Local'
];

const initialDepartments: Department[] = baseDepartments.map((name, index) => ({
  id: `dep-${index + 1}`, name, active: true, areas: []
}));

const seed = <T,>(key: string, fallback: T): T => {
  try {
    const saved = localStorage.getItem(key);
    if (saved) return JSON.parse(saved) as T;
    const backup = localStorage.getItem(`${key}-backup`);
    return backup ? JSON.parse(backup) as T : fallback;
  } catch {
    try {
      const backup = localStorage.getItem(`${key}-backup`);
      return backup ? JSON.parse(backup) as T : fallback;
    } catch { return fallback; }
  }
};


function persist<T>(key: string, value: T) {
  const serialized = JSON.stringify(value);
  const previous = localStorage.getItem(key);
  if (previous) localStorage.setItem(`${key}-backup`, previous);
  localStorage.setItem(key, serialized);
}
function parseWorkbook(buffer: ArrayBuffer): ImportedRow[] {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
  const header = (value: unknown) => String(value).trim().toUpperCase();
  let currentDepartment = '';
  let currentArea = '';
  return rows.map(row => {
    const values = Object.entries(row);
    const get = (names: string[], fallback: number) => {
      const entry = values.find(([key]) => names.includes(header(key)));
      return String(entry?.[1] ?? values[fallback]?.[1] ?? '').trim();
    };
    const department = get(['DEPARTAMENTO', 'DEPARTAMENTO/DEPENDENCIA'], 0);
    const area = get(['AREA', 'ÁREA'], 1);
    if (department) { currentDepartment = department; currentArea = ''; }
    if (area) currentArea = area;
    return { department: currentDepartment, area: currentArea, position: get(['CARGO', 'CARGOS', 'PUESTO'], 2), person: get(['NOMBRES', 'NOMBRE', 'PERSONA', 'RESPONSABLE'], 3) };
  }).filter(row => row.department || row.area || row.position || row.person);
}

function clean(value: string) { return value.trim().replace(/\s+/g, ' ').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase(); }

function normalizeDepartments(source: Department[]) {
  const official = source.filter(department => baseDepartments.includes(department.name)).map(department => ({
    ...department,
    areas: department.areas.filter(area => clean(area.name) !== 'mision juvenil' || department.name === 'Conquistadores Pentecostales')
  }));
  const juvenileAreas = source.flatMap(department => department.areas.filter(area => clean(area.name) === 'mision juvenil'));
  const target = official.find(department => department.name === 'Conquistadores Pentecostales');
  if (target && juvenileAreas.length && !target.areas.some(area => clean(area.name) === 'mision juvenil')) {
    target.areas.push({ ...juvenileAreas[0], id: `area-mision-juvenil-${target.id}` });
  }
  return official;
}

function buildDepartmentsFromRows(rows: ImportedRow[], existing: Department[], people: Person[], keepAssignments: boolean) {
  const personIds = new Map(people.map(person => [clean(person.fullName), person.id]));
  const newPeople: Person[] = [];
  const positionNames = new Set(rows.map(row => clean(row.position)).filter(Boolean));
  rows.forEach(row => {
    if (!row.person || positionNames.has(clean(row.person)) || personIds.has(clean(row.person))) return;
    const id = `person-import-${personIds.size + newPeople.length + 1}`;
    personIds.set(clean(row.person), id);
    newPeople.push({ id, fullName: row.person, active: true });
  });
  const existingOfficial = normalizeDepartments(existing);
  const departments = baseDepartments.map((name, departmentIndex) => {
    const old = existingOfficial.find(item => clean(item.name) === clean(name));
    return { id: old?.id || `dep-${departmentIndex + 1}`, name, active: true, areas: [] as Area[] };
  });
  const reusable = new Map<string, Position[]>();
  existingOfficial.forEach(department => department.areas.forEach(area => area.positions.forEach(position => {
    const key = `${clean(department.name)}|${clean(area.name)}|${clean(position.name)}`;
    reusable.set(key, [...(reusable.get(key) || []), position]);
  })));
  rows.forEach((row, rowIndex) => {
    const department = departments.find(item => clean(item.name) === clean(row.department));
    if (!department || !row.position) return;
    if (clean(row.area) === 'mision juvenil' && clean(department.name) !== 'conquistadores pentecostales') return;
    const areaName = row.area || 'General';
    let area = department.areas.find(item => clean(item.name) === clean(areaName));
    if (!area) {
      area = { id: `area-${department.id}-${department.areas.length + 1}`, name: areaName, active: true, positions: [] };
      department.areas.push(area);
    }
    const key = `${clean(department.name)}|${clean(area.name)}|${clean(row.position)}`;
    const oldPosition = reusable.get(key)?.shift();
    area.positions.push({
      id: oldPosition?.id || `position-${department.id}-${area.id}-${rowIndex + 1}`,
      name: row.position,
      active: true,
      personId: row.person ? personIds.get(clean(row.person)) : keepAssignments ? oldPosition?.personId : undefined
    });
  });
  return { departments, newPeople };
}

function rowsFromDepartments(departments: Department[], people: Person[]): ImportedRow[] {
  return departments.flatMap(department => department.areas.flatMap(area => area.positions.map(position => ({
    department: department.name,
    area: area.name,
    position: position.name,
    person: position.personId ? people.find(person => person.id === position.personId)?.fullName || '' : ''
  }))));
}

function getConflictingAssignments(departments: Department[], people: Person[]) {
  const personNames = new Map(people.map(person => [person.id, clean(person.fullName)]));
  const assignments = new Map<string, { departmentId: string; areaId: string; positionId: string }[]>();
  departments.forEach(department => department.areas.forEach(area => area.positions.forEach(position => {
    if (!position.personId) return;
    const personKey = personNames.get(position.personId) || position.personId;
    const personAssignments = assignments.get(personKey) || [];
    personAssignments.push({ departmentId: department.id, areaId: area.id, positionId: position.id });
    assignments.set(personKey, personAssignments);
  })));

  const conflicts = new Map<string, string>();
  assignments.forEach((personAssignments, personId) => {
    const byArea = new Map<string, typeof personAssignments>();
    personAssignments.forEach(assignment => {
      const areaKey = `${assignment.departmentId}|${assignment.areaId}`;
      byArea.set(areaKey, [...(byArea.get(areaKey) || []), assignment]);
    });
    byArea.forEach(areaAssignments => {
      if (areaAssignments.length < 2) return;
      const { departmentId, areaId } = areaAssignments[0];
      const signature = `${personId}|${departmentId}|${areaId}|${areaAssignments.map(assignment => assignment.positionId).sort().join('|')}`;
      areaAssignments.forEach(assignment => conflicts.set(assignment.positionId, signature));
    });
  });
  return conflicts;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url; anchor.download = filename; anchor.click();
  URL.revokeObjectURL(url);
}

async function imageAsDataUrl(path: string) {
  const response = await fetch(path);
  const blob = await response.blob();
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function App() {
  const [view, setView] = useState<View>('dashboard');
  const [departments, setDepartments] = useState<Department[]>(() => normalizeDepartments(seed('tm-departments', initialDepartments)));
  const [storedPeople, setPeople] = useState<Person[]>(() => seed('tm-people', []));
  const [approvedConflicts, setApprovedConflicts] = useState<string[]>(() => seed('tm-approved-conflicts', []));
  const [selectedDepartment, setSelectedDepartment] = useState<string>(departments[0]?.id || '');
  const [query, setQuery] = useState('');
  const [modal, setModal] = useState<'department' | 'area' | 'position' | 'person' | null>(null);
  const [notice, setNotice] = useState('');
  const treeViewRef = useRef<HTMLDivElement>(null);

  useEffect(() => persist('tm-departments', departments), [departments]);
  const conflictingAssignments = useMemo(() => getConflictingAssignments(departments, storedPeople), [departments, storedPeople]);
  const displayDepartments = useMemo(() => departments.map(department => ({
    ...department,
    areas: department.areas.map(area => ({
      ...area,
      positions: area.positions.map(position => {
        const signature = conflictingAssignments.get(position.id);
        return { ...position, assignmentConflict: !!signature && !approvedConflicts.includes(signature) };
      })
    }))
  })), [departments, conflictingAssignments, approvedConflicts]);
  const conflictingPersonIds = useMemo(() => new Set(displayDepartments.flatMap(department => department.areas.flatMap(area => area.positions
    .filter(position => position.assignmentConflict && position.personId)
    .map(position => position.personId!)))), [displayDepartments]);
  const people = useMemo(() => storedPeople.map(person => ({ ...person, assignmentConflict: conflictingPersonIds.has(person.id) })), [storedPeople, conflictingPersonIds]);

  useEffect(() => persist('tm-people', storedPeople), [storedPeople]);
  useEffect(() => persist('tm-approved-conflicts', approvedConflicts), [approvedConflicts]);
  useEffect(() => {
    navigator.storage?.persist?.().catch(() => undefined);
  }, []);
  useEffect(() => {
    if (localStorage.getItem('tm-catalog-version') === 'base-ods-v3') return;
    fetch('/base.ods').catch(() => fetch('/departamentos.ods'))
      .then(response => response.arrayBuffer())
      .then(buffer => { importRows(parseWorkbook(buffer)); localStorage.setItem('tm-catalog-version', 'base-ods-v3'); })
      .catch(() => setNotice('El archivo de referencia no está disponible para importar'));
  }, [departments]);

  const allPositions = useMemo(() => departments.flatMap(d => d.areas.flatMap(a => a.positions)), [departments]);
  useEffect(() => {
    const positionNames = new Set(allPositions.map(position => clean(position.name)));
    const areaNames = new Set(departments.flatMap(department => department.areas.map(area => clean(area.name))));
    const invalidDirectoryNames = new Set([...positionNames, ...areaNames]);
    const validPeople = people.filter(person => !invalidDirectoryNames.has(clean(person.fullName)));
    if (validPeople.length !== people.length) setPeople(validPeople);
  }, [allPositions, departments, people]);
  const assignedPeopleIds = new Set(allPositions.flatMap(p => p.personId ? [p.personId] : []));
  const stats = {
    departments: departments.filter(d => d.active).length,
    areas: departments.reduce((sum, d) => sum + d.areas.filter(a => a.active).length, 0),
    positions: allPositions.filter(p => p.active).length,
    occupied: allPositions.filter(p => p.active && p.personId).length,
    vacancies: allPositions.filter(p => p.active && !p.personId).length,
    people: people.filter(p => p.active).length,
    unassigned: people.filter(p => p.active && !assignedPeopleIds.has(p.id)).length
  };

  const filteredDepartments = departments.filter(d => d.name.toLowerCase().includes(query.toLowerCase()) || d.areas.some(a => a.name.toLowerCase().includes(query.toLowerCase()) || a.positions.some(p => p.name.toLowerCase().includes(query.toLowerCase()))));
  const organizationDepartments = [
    ...normalizeDepartments(departments),
    ...departments.filter(department => !baseDepartments.includes(department.name) && !department.id.startsWith('dep-import-'))
  ];
  const filteredOrganizationDepartments = organizationDepartments.filter(d => d.name.toLowerCase().includes(query.toLowerCase()) || d.areas.some(a => a.name.toLowerCase().includes(query.toLowerCase()) || a.positions.some(p => p.name.toLowerCase().includes(query.toLowerCase()))));
  const selected = organizationDepartments.find(d => d.id === selectedDepartment) || organizationDepartments[0];

  function addEntity(value: string, type: 'department' | 'area' | 'position' | 'person', parentId?: string) {
    if (!value.trim()) return;
    const id = `${type}-${Date.now()}`;
    if (type === 'department') setDepartments(current => [...current, { id, name: value.trim(), active: true, areas: [] }]);
    if (type === 'area' && parentId) setDepartments(current => current.map(d => d.id === parentId ? { ...d, areas: [...d.areas, { id, name: value.trim(), active: true, positions: [] }] } : d));
    if (type === 'position' && parentId) setDepartments(current => current.map(d => ({ ...d, areas: d.areas.map(a => a.id === parentId ? { ...a, positions: [...a.positions, { id, name: value.trim(), active: true }] } : a) })));
    if (type === 'person') setPeople(current => [...current, { id, fullName: value.trim(), active: true }]);
    setModal(null); setNotice(`${type === 'person' ? 'Persona' : type === 'position' ? 'Cargo' : type === 'area' ? 'Área' : 'Departamento'} creado correctamente`);
    window.setTimeout(() => setNotice(''), 2500);
  }

  function assign(positionId: string, personId: string) {
    const assignedPosition = departments.flatMap(department => department.areas.flatMap(area => area.positions)).find(position => position.id === positionId);
    const conflictSignature = conflictingAssignments.get(positionId);
    if (assignedPosition?.personId === personId && conflictSignature && !approvedConflicts.includes(conflictSignature)) {
      const personName = storedPeople.find(person => person.id === personId)?.fullName || 'Esta persona';
      const conflictingPositions = departments.flatMap(department => department.areas.flatMap(area => area.positions
        .filter(position => conflictingAssignments.get(position.id) === conflictSignature)
        .map(position => `${department.name} / ${area.name}: ${position.name}`)));
      if (window.confirm(`${personName} está asignada a varios cargos:\n\n${conflictingPositions.join('\n')}\n\n¿Deseas mantener estas asignaciones?`)) {
        setApprovedConflicts(current => current.includes(conflictSignature) ? current : [...current, conflictSignature]);
      }
      return;
    }
    setDepartments(current => current.map(d => ({ ...d, areas: d.areas.map(a => ({ ...a, positions: a.positions.map(p => p.id === positionId ? { ...p, personId } : p) })) })));
    setNotice('Asignación guardada'); window.setTimeout(() => setNotice(''), 2500);
  }

  function updatePositionPerson(positionId: string, fullName: string) {
    const name = fullName.trim();
    setDepartments(current => current.map(department => ({ ...department, areas: department.areas.map(area => ({ ...area, positions: area.positions.map(position => {
      if (position.id !== positionId) return position;
      if (!name) return { ...position, personId: undefined };
      const existing = people.find(person => clean(person.fullName) === clean(name));
      const personId = existing?.id || `person-${Date.now()}`;
      if (!existing) setPeople(currentPeople => [...currentPeople, { id: personId, fullName: name, active: true }]);
      return { ...position, personId };
    }) })) })));
  }

  function deleteEntity(type: 'department' | 'area' | 'position', id: string) {
    if (type === 'department') setDepartments(current => current.filter(department => department.id !== id));
    if (type === 'area') setDepartments(current => current.map(department => ({ ...department, areas: department.areas.filter(area => area.id !== id) })));
    if (type === 'position') setDepartments(current => current.map(department => ({ ...department, areas: department.areas.map(area => ({ ...area, positions: area.positions.filter(position => position.id !== id) })) })));
    setNotice(`${type === 'department' ? 'Departamento' : type === 'area' ? 'Área' : 'Cargo'} eliminado`);
    window.setTimeout(() => setNotice(''), 2500);
  }

  function deletePerson(personId: string) {
    setPeople(current => current.filter(person => person.id !== personId));
    setDepartments(current => current.map(department => ({ ...department, areas: department.areas.map(area => ({ ...area, positions: area.positions.map(position => position.personId === personId ? { ...position, personId: undefined } : position) })) })));
    setNotice('Persona eliminada y cargo liberado');
    window.setTimeout(() => setNotice(''), 2500);
  }

  function importRows(rows: ImportedRow[]) {
    const result = buildDepartmentsFromRows(rows, departments, people, true);
    setDepartments(result.departments);
    if (result.newPeople.length) setPeople(current => [...current, ...result.newPeople]);
    localStorage.setItem('tm-reference-imported', 'true');
    setSelectedDepartment(result.departments[0]?.id || '');
    setNotice(`Importación completada: ${rows.length} filas, ${result.newPeople.length} personas nuevas`);
    window.setTimeout(() => setNotice(''), 4000);
  }

  function importFile(file: File) {
    file.arrayBuffer().then(buffer => importRows(parseWorkbook(buffer))).catch(() => setNotice('No se pudo leer el archivo seleccionado'));
  }

  function resetProcess() {
    if (!window.confirm('Se borrarán únicamente las personas asignadas a los cargos. Los cargos, departamentos y áreas se conservarán. ¿Continuar?')) return;
    setPeople([]);
    fetch('/base.ods').catch(() => fetch('/departamentos.ods'))
      .then(response => response.arrayBuffer())
      .then(buffer => {
        const result = buildDepartmentsFromRows(parseWorkbook(buffer), departments, [], false);
        setDepartments(result.departments);
        localStorage.setItem('tm-catalog-version', 'base-ods-v3');
        setNotice('Proceso reiniciado: cargos conservados y asignaciones eliminadas');
      })
      .catch(() => setNotice('No se pudo restaurar el catálogo de cargos'));
    window.setTimeout(() => setNotice(''), 3500);
  }

  async function exportExcel() {
    const rows = rowsFromDepartments(normalizeDepartments(departments), people);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Gestión de Directivas';
    const worksheet = workbook.addWorksheet('Directivas 2027');
    worksheet.columns = [{ header: 'DEPARTAMENTO', key: 'department', width: 34 }, { header: 'AREA', key: 'area', width: 36 }, { header: 'CARGO', key: 'position', width: 52 }, { header: 'NOMBRES', key: 'person', width: 34 }, { header: '', key: 'spacer', width: 4 }, { header: '', key: 'letterhead', width: 25 }];
    rows.forEach(row => worksheet.addRow({ department: row.department, area: row.area, position: row.position, person: row.person }));
    worksheet.autoFilter = { from: 'A1', to: `D${rows.length + 1}` };
    worksheet.getRow(1).height = 28;
    worksheet.getRow(1).eachCell(cell => { cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF06487F' } }; cell.alignment = { vertical: 'middle' }; cell.border = { bottom: { style: 'thin', color: { argb: 'FFB9CFE0' } } }; });
    worksheet.eachRow((row, rowNumber) => { if (rowNumber > 1) { row.eachCell({ includeEmpty: true }, cell => { cell.border = { bottom: { style: 'hair', color: { argb: 'FFD9E3EC' } } }; cell.alignment = { vertical: 'top', wrapText: true }; }); } });
    worksheet.mergeCells('F2:J2'); worksheet.getCell('F2').value = 'IPUC Riohacha - Central'; worksheet.getCell('F2').font = { bold: true, size: 16, color: { argb: 'FF06487F' } };
    worksheet.mergeCells('F3:J3'); worksheet.getCell('F3').value = 'Directivas 2027'; worksheet.getCell('F3').font = { bold: true, size: 12, color: { argb: 'FFFFA900' } };
    const logoData = await imageAsDataUrl('/IPUC_COLOR para fondo claro.png').catch(() => '');
    if (logoData) { const imageId = workbook.addImage({ base64: logoData, extension: 'png' }); worksheet.addImage(imageId, { tl: { col: 5, row: 0 }, ext: { width: 105, height: 66 } }); }
    const output = await workbook.xlsx.writeBuffer();
    downloadBlob(new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'IPUC-Riohacha-Central-Directivas-2027.xlsx');
  }

  async function exportPdfColumns() {
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    try { pdf.addImage(await imageAsDataUrl('/IPUC_COLOR para fondo claro.png'), 'PNG', 14, 10, 24, 18); } catch { /* logo opcional */ }
    pdf.setFontSize(16); pdf.setTextColor(25, 59, 57); pdf.text('IPUC Riohacha - Central', 44, 17);
    pdf.setFontSize(11); pdf.text('Directivas 2027', 44, 24);
    let y = 36; const columns = [14, 78, 145, 218];
    pdf.setFillColor(25, 59, 57); pdf.rect(14, y - 6, 269, 9, 'F'); pdf.setTextColor(255, 255, 255); pdf.setFontSize(9);
    ['Departamento', 'Área', 'Cargo', 'Persona'].forEach((label, index) => pdf.text(label, columns[index], y));
    pdf.setTextColor(29, 43, 42); y += 9;
    rowsFromDepartments(departments, people).forEach(row => { if (y > 190) { pdf.addPage(); y = 18; } pdf.setDrawColor(220, 232, 226); pdf.line(14, y + 2, 283, y + 2); pdf.text(row.department.slice(0, 30), columns[0], y); pdf.text(row.area.slice(0, 32), columns[1], y); pdf.text(row.position.slice(0, 35), columns[2], y); pdf.text(row.person.slice(0, 28), columns[3], y); y += 8; });
    pdf.save('IPUC-Riohacha-Central-Directivas-2027.pdf');
  }

  async function exportPdfTree() {
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    try { pdf.addImage(await imageAsDataUrl('/IPUC_COLOR para fondo claro.png'), 'PNG', 14, 9, 22, 17); } catch { /* logo opcional */ }
    pdf.setTextColor(6, 72, 127); pdf.setFontSize(15); pdf.text('IPUC Riohacha - Central', 42, 16); pdf.setFontSize(10); pdf.setTextColor(255, 169, 0); pdf.text('Directivas 2027 | Árbol organizacional', 42, 23);
    let y = 34;
    const ensureSpace = (height: number) => { if (y + height > 195) { pdf.addPage(); y = 18; } };
    const text = (value: string, x: number, size = 8, color = [24, 52, 76] as [number, number, number]) => { pdf.setFontSize(size); pdf.setTextColor(...color); pdf.text(value, x, y); };
    pdf.setDrawColor(6, 72, 127); pdf.setLineWidth(0.7); pdf.line(148, y - 7, 148, y - 2); ensureSpace(14); pdf.setFillColor(232, 242, 251); pdf.roundedRect(115, y - 8, 66, 11, 2, 2, 'F'); pdf.setFontSize(10); pdf.setTextColor(6, 72, 127); pdf.text('IPUC Riohacha - Central', 121, y - 1); y += 12;
    normalizeDepartments(departments).forEach(department => { ensureSpace(18); pdf.setFillColor(6, 72, 127); pdf.roundedRect(18, y - 7, 95, 9, 2, 2, 'F'); pdf.setFontSize(9); pdf.setTextColor(255, 255, 255); pdf.text(department.name, 22, y - 1); pdf.setDrawColor(133, 181, 214); pdf.line(113, y - 3, 127, y - 3); y += 8; department.areas.forEach(area => { ensureSpace(14); pdf.setDrawColor(176, 207, 226); pdf.line(127, y - 5, 127, y + 2); pdf.setFillColor(232, 242, 251); pdf.roundedRect(132, y - 7, 74, 8, 1.5, 1.5, 'F'); pdf.setFontSize(8); pdf.setTextColor(6, 72, 127); pdf.text(area.name.slice(0, 42), 136, y - 1); pdf.line(206, y - 3, 218, y - 3); y += 7; area.positions.forEach(position => { ensureSpace(9); const person = position.personId ? people.find(item => item.id === position.personId)?.fullName || 'Sin nombre' : 'Vacante'; pdf.setDrawColor(220, 230, 237); pdf.line(218, y - 4, 218, y + 1); pdf.setFontSize(7.5); pdf.setTextColor(24, 52, 76); pdf.text(`${position.name.slice(0, 34)} | ${person.slice(0, 25)}`, 222, y); y += 7; }); y += 2; }); y += 4; });
    pdf.save('IPUC-Riohacha-Central-Arbol-2027.pdf');
  }

  async function exportPdfTreeView() {
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const normalizedDepartments = normalizeDepartments(departments);
    const logoData = await imageAsDataUrl('/IPUC_COLOR para fondo claro.png').catch(() => '');
    const page = { width: 297, height: 210, left: 12, right: 285, bottom: 194 };
    const drawHeader = () => {
      pdf.setFillColor(255, 255, 255); pdf.rect(0, 0, page.width, page.height, 'F');
      pdf.setTextColor(6, 72, 127); pdf.setFontSize(15); pdf.text('IPUC Riohacha - Central', 37, 13);
      pdf.setTextColor(210, 143, 0); pdf.setFontSize(9); pdf.text('Directivas 2027 | Vista del árbol', 37, 20);
      pdf.setDrawColor(190, 210, 224); pdf.line(page.left, 25, page.right, 25);
    };
    const wrap = (value: string, width: number) => pdf.splitTextToSize(value || 'Vacante', width) as string[];
    let y = 34;
    drawHeader();
    if (logoData) pdf.addImage(logoData, 'PNG', 14, 5, 17, 17);
    normalizedDepartments.forEach(department => {
      if (y > page.bottom - 18) { pdf.addPage(); drawHeader(); if (logoData) pdf.addImage(logoData, 'PNG', 14, 5, 17, 17); y = 34; }
      pdf.setFillColor(6, 72, 127); pdf.roundedRect(15, y - 7, 88, 10, 2, 2, 'F'); pdf.setTextColor(255, 255, 255); pdf.setFontSize(9); pdf.text(department.name, 19, y - 1); y += 8;
      department.areas.forEach(area => {
        if (y > page.bottom - 24) { pdf.addPage(); drawHeader(); if (logoData) pdf.addImage(logoData, 'PNG', 14, 5, 17, 17); y = 34; }
        pdf.setDrawColor(178, 207, 225); pdf.line(31, y - 5, 31, y + 2); pdf.setFillColor(232, 242, 251); pdf.roundedRect(36, y - 7, 72, 9, 2, 2, 'F'); pdf.setTextColor(6, 72, 127); pdf.setFontSize(8); pdf.text(area.name, 40, y - 1); y += 7;
        let x = 44;
        area.positions.forEach(position => {
          const person = position.personId ? people.find(item => item.id === position.personId)?.fullName || 'Sin nombre' : 'Vacante';
          const positionLines = wrap(position.name, 48);
          const personLines = wrap(person, 48);
          const cardHeight = Math.max(16, 6 + positionLines.length * 4 + personLines.length * 4);
          if (x + 58 > page.right) { x = 44; y += 19; }
          if (y + cardHeight > page.bottom) { pdf.addPage(); drawHeader(); if (logoData) pdf.addImage(logoData, 'PNG', 14, 5, 17, 17); y = 34; x = 44; }
          pdf.setDrawColor(201, 216, 226); pdf.setFillColor(255, 255, 255); pdf.roundedRect(x, y - 5, 55, cardHeight, 2, 2, 'FD');
          pdf.setTextColor(24, 52, 76); pdf.setFontSize(7); pdf.text(positionLines, x + 3, y + 1, { maxWidth: 49 });
          pdf.setTextColor(49, 94, 86); pdf.setFontSize(6.5); pdf.text(personLines, x + 3, y + 1 + positionLines.length * 4, { maxWidth: 49 });
          x += 59;
        });
        y += 19;
      });
      y += 8;
    });
    pdf.save('IPUC-Riohacha-Central-Arbol-Vista-2027.pdf');
  }

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><img className="brand-logo" src="/IPUC_COLOR para fondo claro.png" alt="Logo IPUC" /><div><strong>Gestión de Directivas</strong><span>IPUC Riohacha - Central</span></div></div>
      <div className="workspace-label">ESPACIO LOCAL <span className="online-dot" /></div>
      <nav className="nav">
        <NavItem icon={<LayoutDashboard size={18} />} label="Dashboard" active={view === 'dashboard'} onClick={() => setView('dashboard')} />
        <NavItem icon={<Building2 size={18} />} label="Organización" active={view === 'organization'} onClick={() => setView('organization')} />
        <NavItem icon={<GitBranch size={18} />} label="Árbol" active={view === 'tree'} onClick={() => setView('tree')} />
        <NavItem icon={<UsersRound size={18} />} label="Personal" active={view === 'people'} onClick={() => setView('people')} />
        <NavItem icon={<BriefcaseBusiness size={18} />} label="Asignaciones" active={view === 'assignments'} onClick={() => setView('assignments')} />
        <NavItem icon={<CircleHelp size={18} />} label="Cargos vacantes" active={view === 'vacancies'} onClick={() => setView('vacancies')} badge={stats.vacancies} />
        <NavItem icon={<UserRound size={18} />} label="Personal sin cargo" active={view === 'unassigned'} onClick={() => setView('unassigned')} />
        <NavItem icon={<BarChart3 size={18} />} label="Estadísticas" active={view === 'dashboard'} onClick={() => setView('dashboard')} />
        <div className="nav-divider" />
        <NavItem icon={<FileUp size={18} />} label="Importar" active={false} onClick={() => setView('settings')} />
        <NavItem icon={<FileDown size={18} />} label="Exportar Excel" active={false} onClick={exportExcel} />
        <NavItem icon={<FileDown size={18} />} label="Exportar PDF" active={false} onClick={exportPdfColumns} />
        <NavItem icon={<History size={18} />} label="Historial" active={view === 'history'} onClick={() => setView('history')} />
      </nav>
      <div className="sidebar-bottom"><NavItem icon={<Settings2 size={18} />} label="Configuración" active={view === 'settings'} onClick={() => setView('settings')} /><div className="local-storage"><Archive size={15} /><span>Datos guardados localmente</span></div></div>
    </aside>
    <main className="main-content">
      <header className="topbar"><button className="mobile-menu" aria-label="Abrir menú"><Menu size={20} /></button><div className="global-search"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Buscar personas, cargos, áreas o departamentos..." /></div><div className="topbar-actions"><button className="icon-button" aria-label="Ayuda"><CircleHelp size={19} /></button><div className="user-chip"><div className="avatar">GD</div><div><strong>Administrador</strong><span>Sesión local</span></div><ChevronDown size={15} /></div></div></header>
      <section className="page">
        {notice && <div className="toast"><Activity size={16} />{notice}<button onClick={() => setNotice('')}><X size={15} /></button></div>}
        {view === 'dashboard' && <Dashboard stats={stats} onNavigate={setView} />}
        {view === 'organization' && <Organization departments={filteredOrganizationDepartments.map(department => displayDepartments.find(item => item.id === department.id) || department)} people={people} selected={displayDepartments.find(item => item.id === selected?.id) || selected} selectedId={selectedDepartment} onSelect={setSelectedDepartment} onAdd={setModal} onDelete={deleteEntity} onAssign={assign} />}
        {view === 'tree' && <OrganizationTree departments={displayDepartments} people={people} treeViewRef={treeViewRef} onAssign={assign} onExportPdfTree={exportPdfTree} onExportPdfTreeView={exportPdfTreeView} onExportPdfColumns={exportPdfColumns} />}
        {view === 'people' && <People people={people} positions={allPositions} onAdd={() => setModal('person')} onDelete={deletePerson} />}
        {view === 'assignments' && <Assignments departments={filteredDepartments} people={people} onAssign={assign} />}
        {view === 'vacancies' && <Vacancies departments={departments} onAssign={() => setView('assignments')} />}
        {view === 'unassigned' && <Unassigned people={people} assignedIds={assignedPeopleIds} onAdd={() => setModal('person')} />}
        {view === 'history' && <EmptyState icon={<History size={30} />} title="Historial local" text="Las asignaciones y cambios aparecerán aquí." />}
        {view === 'settings' && <Settings onImport={importFile} onExportExcel={exportExcel} onExportPdf={exportPdfColumns} onReset={resetProcess} />}
      </section>
    </main>
    {modal && <EntityModal type={modal} departments={departments} onClose={() => setModal(null)} onSave={addEntity} />}
  </div>;
}

function NavItem({ icon, label, active, onClick, badge }: { icon: React.ReactNode; label: string; active: boolean; onClick: () => void; badge?: number }) { return <button className={`nav-item ${active ? 'active' : ''}`} onClick={onClick}>{icon}<span>{label}</span>{badge ? <b>{badge}</b> : null}</button>; }
function Header({ eyebrow, title, text, action }: { eyebrow: string; title: string; text: string; action?: React.ReactNode }) { const displayTitle = title === 'Buenos días, Administrador' ? 'Buenos días, Derick' : title; return <div className="page-header"><div><div className="eyebrow">{eyebrow}</div><h1>{displayTitle}</h1><p>{text}</p></div>{action}</div>; }
function Dashboard({ stats, onNavigate }: { stats: Record<string, number>; onNavigate: (v: View) => void }) { return <><Header eyebrow="Resumen operativo" title="Buenos días, Administrador" text="Aquí tienes una vista clara de la estructura y el estado actual de tu organización." action={<button className="primary-button" onClick={() => onNavigate('assignments')}><BriefcaseBusiness size={17} /> Gestionar asignaciones</button>} /><div className="stats-grid">{[['departments','Departamentos','Catálogo oficial'],['areas','Áreas','Estructura activa'],['positions','Cargos','Total registrados'],['occupied','Ocupados','Asignados actualmente'],['vacancies','Vacantes','Requieren asignación'],['people','Personas','Personal activo']].map(([key, label, caption]) => <div className="stat-card" key={key}><div className="stat-top"><span>{label}</span><ArrowUpRight size={16} /></div><strong>{stats[key]}</strong><small>{caption}</small></div>)}</div><div className="dashboard-grid"><section className="panel activity-panel"><div className="panel-heading"><div><div className="eyebrow">Actividad</div><h2>Estado de asignaciones</h2></div><button className="quiet-button" onClick={() => onNavigate('assignments')}>Ver detalle <ArrowUpRight size={15} /></button></div><div className="empty-chart"><div className="chart-ring" style={{ '--progress': stats.positions ? `${(stats.occupied / stats.positions) * 100}%` : '0%' } as React.CSSProperties}><strong>{stats.positions ? Math.round((stats.occupied / stats.positions) * 100) : 0}%</strong><span>ocupación</span></div><div className="legend"><span><i className="dot green" /> Cargos ocupados <b>{stats.occupied}</b></span><span><i className="dot red" /> Cargos vacantes <b>{stats.vacancies}</b></span></div></div></section><section className="panel quick-panel"><div className="panel-heading"><div><div className="eyebrow">Acciones rápidas</div><h2>Continúa tu trabajo</h2></div></div><QuickAction icon={<Building2 size={19} />} text="Configurar organización" onClick={() => onNavigate('organization')} /><QuickAction icon={<UserRound size={19} />} text="Registrar una persona" onClick={() => onNavigate('people')} /><QuickAction icon={<BriefcaseBusiness size={19} />} text="Revisar cargos vacantes" onClick={() => onNavigate('vacancies')} /></section></div><div className="info-banner"><div className="banner-icon"><Archive size={20} /></div><div><strong>Todo está guardado en este equipo</strong><p>T Maestro funciona sin conexión. Tus datos permanecen en el almacenamiento local de esta aplicación.</p></div><button className="quiet-button" onClick={() => onNavigate('settings')}>Configuración <ArrowUpRight size={15} /></button></div></>; }
function QuickAction({ icon, text, onClick }: { icon: React.ReactNode; text: string; onClick: () => void }) { return <button className="quick-action" onClick={onClick}><span>{icon}</span>{text}<ArrowUpRight size={16} /></button>; }
function Organization({ departments, people, selected, selectedId, onSelect, onAdd, onDelete, onAssign }: { departments: Department[]; people: Person[]; selected?: Department; selectedId: string; onSelect: (id: string) => void; onAdd: (type: 'department' | 'area' | 'position' | 'person') => void; onDelete: (type: 'department' | 'area' | 'position', id: string) => void; onAssign: (positionId: string, personId: string) => void }) {
  return <><Header eyebrow="Estructura organizacional" title="Organización" text="Administra departamentos, áreas y cargos desde una sola vista." action={<button className="primary-button" onClick={() => onAdd('department')}><Plus size={17} /> Nuevo departamento</button>} /><div className="org-layout"><section className="panel org-tree"><div className="panel-heading"><div><h2>Departamentos</h2><p>{departments.length} departamentos oficiales</p></div><Building2 size={20} className="muted-icon" /></div><div className="tree-list">{departments.map(d => <div className={`tree-department-row ${selectedId === d.id ? 'selected' : ''}`} key={d.id}><button className="tree-department" onClick={() => onSelect(d.id)}><span className="tree-dot" />{d.name}<span className="tree-count">{d.areas.reduce((n, a) => n + a.positions.length, 0)}</span></button><button className="delete-button" title="Eliminar departamento" onClick={() => onDelete('department', d.id)}><X size={15} /></button></div>)}</div></section><section className="panel org-detail"><div className="panel-heading"><div><div className="eyebrow">Departamento seleccionado</div><h2>{selected?.name || 'Sin departamento'}</h2><p>{selected?.areas.length || 0} áreas · {selected?.areas.reduce((n, a) => n + a.positions.length, 0) || 0} cargos</p></div><div className="header-actions"><button className="secondary-button" disabled={!selected} onClick={() => onAdd('area')}><Plus size={16} /> Nueva área</button><button className="delete-button" title="Eliminar departamento" disabled={!selected} onClick={() => selected && onDelete('department', selected.id)}><X size={17} /></button></div></div>{selected?.areas.length ? <><div className="org-table-header"><span>Cargo</span><span>Persona que ocupa el cargo</span><span>Estado</span></div>{selected.areas.map(area => <div className="area-block" key={area.id}><div className="area-heading"><span><Building2 size={16} />{area.name}</span><div className="area-actions"><button className="text-button" onClick={() => onAdd('position')}><Plus size={14} /> Cargo</button><button className="delete-button" title="Eliminar área" onClick={() => onDelete('area', area.id)}><X size={15} /></button></div></div>{area.positions.length ? area.positions.map(position => <div className="position-row" key={position.id}><BriefcaseBusiness size={16} /><span>{position.name}</span><PositionPersonPicker people={people} personId={position.personId} onAssign={personId => onAssign(position.id, personId)} /><StatusBadge occupied={!!position.personId} /><button className="delete-button" title="Eliminar cargo" onClick={() => onDelete('position', position.id)}><X size={15} /></button></div>) : <div className="mini-empty">Sin cargos en esta área</div>}</div>)}</> : <EmptyState icon={<Building2 size={30} />} title="Empieza a construir la estructura" text="Añade un área para organizar los cargos de este departamento." action={<button className="secondary-button" onClick={() => onAdd('area')}><Plus size={16} /> Nueva área</button>} />}</section></div></>;
}

function PositionPersonPicker({ people, personId, assignmentConflict = people.find(person => person.id === personId)?.assignmentConflict || false, onAssign }: { people: Person[]; personId?: string; assignmentConflict?: boolean; onAssign: (personId: string) => void }) {
  const [search, setSearch] = useState('');
  const [open, setOpen] = useState(false);
  const pickerRef = useRef<HTMLDivElement>(null);
  const current = people.find(person => person.id === personId)?.fullName || '';
  const results = people.filter(person => person.fullName.toLocaleLowerCase().includes(search.toLocaleLowerCase())).slice(0, 8);
  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      if (pickerRef.current && !pickerRef.current.contains(event.target as Node)) {
        setOpen(false);
        setSearch('');
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, []);
  return <div className="person-picker-inline" ref={pickerRef}><div className={`person-picker-input ${assignmentConflict ? 'assignment-conflict' : ''}`}><Search size={14} /><input value={open ? search : current} placeholder="Buscar persona..." onFocus={() => { setOpen(true); setSearch(''); }} onChange={event => { setOpen(true); setSearch(event.target.value); }} /></div>{assignmentConflict && personId && <button type="button" className="confirm-assignment-button" title="Confirmar que se mantengan los cargos" aria-label="Confirmar que se mantengan los cargos" onClick={() => onAssign(personId)}><ShieldCheck size={15} /></button>}{open && <div className="person-picker-results">{results.length ? results.map(person => <button type="button" key={person.id} onMouseDown={event => event.preventDefault()} onClick={() => { onAssign(person.id); setOpen(false); setSearch(''); }}><UserRound size={14} />{person.fullName}</button>) : <span>No hay personas coincidentes</span>}{current && <button type="button" className="clear-person" onClick={() => { onAssign(''); setOpen(false); }}>Quitar asignación</button>}</div>}</div>;
}

function OrganizationTree({ departments, people, treeViewRef, onAssign, onExportPdfTree, onExportPdfTreeView, onExportPdfColumns }: { departments: Department[]; people: Person[]; treeViewRef: React.RefObject<HTMLDivElement | null>; onAssign: (positionId: string, personId: string) => void; onExportPdfTree: () => void; onExportPdfTreeView: () => void; onExportPdfColumns: () => void }) {
  return <><Header eyebrow="Vista general" title="Árbol organizacional" text="Consulta departamentos, áreas, cargos y responsables en una sola vista." action={<div className="header-actions"><button className="primary-button" onClick={onExportPdfTreeView}><GitBranch size={17} /> Exportar vista</button><button className="secondary-button" onClick={onExportPdfTree}><GitBranch size={17} /> Exportar árbol</button><button className="secondary-button" onClick={onExportPdfColumns}><FileDown size={17} /> Exportar columnas</button></div>} /><section className="panel organigram-panel" ref={treeViewRef}><div className="org-root">IPUC Riohacha - Central<br /><small>Directivas 2027</small></div>{departments.map(department => <div className="org-branch" key={department.id}><div className="org-node department-node">{department.name}</div>{department.areas.map(area => <div className="org-area" key={area.id}><div className="org-node area-node">{area.name}</div>{area.positions.map(position => {
    const person = people.find(item => item.id === position.personId);
    return <div className="org-position" key={position.id}><BriefcaseBusiness size={14} /> <span><strong>{position.name}</strong><small className={position.assignmentConflict ? 'assignment-conflict' : undefined} style={position.assignmentConflict ? { color: '#b42318', fontWeight: 700 } : undefined}>{person?.fullName || (position.personId ? 'Sin nombre' : 'Vacante')}</small></span>{position.assignmentConflict && position.personId && <button type="button" className="confirm-assignment-button" title="Confirmar que se mantengan los cargos" aria-label="Confirmar que se mantengan los cargos" onClick={() => onAssign(position.id, position.personId!)}><ShieldCheck size={15} /></button>}</div>;
  })}</div>)}</div>)}</section></>;
}
function StatusBadge({ occupied }: { occupied: boolean }) { return <span className={`status ${occupied ? 'occupied' : 'vacant'}`}><i />{occupied ? 'Ocupado' : 'Vacante'}</span>; }
function People({ people, positions, onAdd, onDelete }: { people: Person[]; positions: Position[]; onAdd: () => void; onDelete: (personId: string) => void }) { const counts = new Map<string, number>(); positions.forEach(p => p.personId && counts.set(p.personId, (counts.get(p.personId) || 0) + 1)); return <><Header eyebrow="Directorio local" title="Personal" text="Consulta y administra las personas disponibles para asignación." action={<button className="primary-button" onClick={onAdd}><Plus size={17} /> Nueva persona</button>} /><section className="panel table-panel"><div className="table-toolbar"><div className="table-title"><UsersRound size={20} /><strong>{people.length} personas registradas</strong></div><div className="filter-chip">Todos <ChevronDown size={14} /></div></div>{people.length ? <table><thead><tr><th>N.º</th><th>Nombre completo</th><th>Identificación</th><th>Cargos</th><th>Estado</th><th /></tr></thead><tbody>{people.map((p, index) => <tr key={p.id}><td className="person-number">{index + 1}</td><td><div className="person-cell"><div className="small-avatar">{p.fullName.slice(0, 2).toUpperCase()}</div><strong>{p.fullName}</strong></div></td><td>{p.identification || 'Sin registrar'}</td><td>{counts.get(p.id) || 0}</td><td><StatusBadge occupied={!!counts.get(p.id)} /></td><td><button className="delete-button" title="Eliminar persona" onClick={() => { if (window.confirm(`¿Eliminar a ${p.fullName}?`)) onDelete(p.id); }}><X size={16} /></button></td></tr>)}</tbody></table> : <EmptyState icon={<UsersRound size={30} />} title="Todavía no hay personas" text="Registra personas manualmente o importa tu archivo de personal." action={<button className="secondary-button" onClick={onAdd}><Plus size={16} /> Registrar persona</button>} />}</section></>; }
type AssignmentPosition = Position & { department: string; area: string };

function Assignments({ departments, people, onAssign }: { departments: Department[]; people: Person[]; onAssign: (positionId: string, personId: string) => void }) { const [selectedPosition, setSelectedPosition] = useState<AssignmentPosition | null>(null); const [search, setSearch] = useState(''); const positions: AssignmentPosition[] = departments.flatMap(d => d.areas.flatMap(a => a.positions.map(p => ({ ...p, department: d.name, area: a.name })))); const results = people.filter(p => p.fullName.toLowerCase().includes(search.toLowerCase())); return <><Header eyebrow="Trabajo principal" title="Asignaciones" text="Selecciona un cargo y asigna manualmente una persona." /><div className="assignment-layout"><section className="panel assignment-tree"><div className="panel-heading"><h2>Cargos</h2><span className="count-label">{positions.length}</span></div><div className="assignment-list">{positions.map(p => <button className={`assignment-item ${selectedPosition?.id === p.id ? 'selected' : ''}`} key={p.id} onClick={() => setSelectedPosition(p)}><span><BriefcaseBusiness size={15} />{p.name}<small>{p.department} · {p.area}</small></span><StatusBadge occupied={!!p.personId} /></button>)}</div>{!positions.length && <div className="mini-empty">Crea cargos en Organización para comenzar.</div>}</section><section className="panel assignment-focus">{selectedPosition ? <><div className="eyebrow">Cargo seleccionado</div><h2>{selectedPosition.name}</h2><p className="muted-copy">{selectedPosition.department} · {selectedPosition.area}</p><div className="focus-status"><StatusBadge occupied={!!selectedPosition.personId} /><p>{selectedPosition.personId ? people.find(p => p.id === selectedPosition.personId)?.fullName : 'Este cargo está disponible para asignación'}</p></div></> : <EmptyState icon={<BriefcaseBusiness size={30} />} title="Selecciona un cargo" text="Elige un cargo del árbol para ver su información y asignar una persona." />}</section><section className="panel people-picker"><div className="panel-heading"><div><h2>Personas</h2><p>Asignación manual</p></div></div><div className="inner-search"><Search size={16} /><input value={search} onChange={e => setSearch(e.target.value)} placeholder="Buscar persona..." /></div>{selectedPosition && results.map(p => <div className="picker-row" key={p.id}><div className="person-cell"><div className="small-avatar">{p.fullName.slice(0, 2).toUpperCase()}</div><div><strong>{p.fullName}</strong><small>{people.filter(x => x.id === p.id).length ? 'Disponible para asignar' : ''}</small></div></div><button className="assign-button" onClick={() => onAssign(selectedPosition.id, p.id)}>Asignar</button></div>)}{!selectedPosition && <div className="mini-empty">Selecciona un cargo para buscar personas.</div>}{selectedPosition && !results.length && <div className="mini-empty">No hay coincidencias.</div>}</section></div></>; }
function Vacancies({ departments, onAssign }: { departments: Department[]; onAssign: () => void }) { const vacancies = departments.flatMap(d => d.areas.flatMap(a => a.positions.filter(p => !p.personId).map(p => ({ ...p, department: d.name, area: a.name })))); return <><Header eyebrow="Seguimiento" title="Cargos vacantes" text="Identifica rápidamente los cargos que aún necesitan una asignación." action={<button className="primary-button" onClick={onAssign}><BriefcaseBusiness size={17} /> Asignar persona</button>} /><section className="panel table-panel"><div className="table-toolbar"><div className="table-title"><span className="red-icon"><BriefcaseBusiness size={18} /></span><strong>{vacancies.length} cargos vacantes</strong></div></div>{vacancies.length ? <table><thead><tr><th>Cargo</th><th>Departamento</th><th>Área</th><th /></tr></thead><tbody>{vacancies.map(v => <tr key={v.id}><td><strong>{v.name}</strong></td><td>{v.department}</td><td>{v.area}</td><td><button className="assign-button" onClick={onAssign}>Asignar</button></td></tr>)}</tbody></table> : <EmptyState icon={<BriefcaseBusiness size={30} />} title="No hay cargos vacantes" text="Crea cargos en Organización para hacer seguimiento a la ocupación." />}</section></>; }
function Unassigned({ people, assignedIds, onAdd }: { people: Person[]; assignedIds: Set<string>; onAdd: () => void }) { const unassigned = people.filter(p => !assignedIds.has(p.id)); return <><Header eyebrow="Seguimiento" title="Personal sin cargo" text="Personas registradas que todavía no tienen una asignación activa." action={<button className="primary-button" onClick={onAdd}><Plus size={17} /> Nueva persona</button>} /><section className="panel table-panel"><div className="table-toolbar"><div className="table-title"><span className="orange-icon"><UserRound size={18} /></span><strong>{unassigned.length} personas sin cargo</strong></div></div>{unassigned.length ? <table><thead><tr><th>Persona</th><th>Estado</th><th /></tr></thead><tbody>{unassigned.map(p => <tr key={p.id}><td><div className="person-cell"><div className="small-avatar">{p.fullName.slice(0, 2).toUpperCase()}</div><strong>{p.fullName}</strong></div></td><td><span className="status warning"><i />Sin cargo</span></td><td><button className="assign-button">Asignar</button></td></tr>)}</tbody></table> : <EmptyState icon={<UserRound size={30} />} title="No hay personas sin cargo" text="Registra personas o importa el archivo de personal." />}</section></>; }
function Settings({ onImport, onExportExcel, onExportPdf, onReset }: { onImport: (file: File) => void; onExportExcel: () => void; onExportPdf: () => void; onReset: () => void }) { return <><Header eyebrow="Preferencias y datos" title="Configuración" text="Controla la información local y el comportamiento de la aplicación." /><div className="settings-grid"><section className="panel settings-card"><div className="settings-icon"><FileUp size={19} /></div><h2>Importar datos</h2><p>Lee archivos ODS, XLSX, XLS o CSV y completa la estructura sin duplicar cargos.</p><label className="secondary-button file-button"><FileUp size={16} /> Seleccionar archivo<input type="file" accept=".ods,.xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) onImport(file); event.currentTarget.value = ''; }} /></label></section><section className="panel settings-card"><div className="settings-icon"><FileDown size={19} /></div><h2>Exportar directivas</h2><p>Genera documentos profesionales con el encabezado institucional y todas las asignaciones.</p><div className="header-actions"><button className="secondary-button" onClick={onExportExcel}><FileDown size={16} /> Excel</button><button className="secondary-button" onClick={onExportPdf}><FileDown size={16} /> PDF</button></div></section><section className="panel settings-card danger-card"><div className="settings-icon"><Archive size={19} /></div><h2>Reiniciar proceso</h2><p>Conserva departamentos, áreas y cargos; elimina únicamente personas y asignaciones.</p><button className="secondary-button" onClick={onReset}><Archive size={16} /> Reiniciar asignaciones</button></section></div></>; }
function EmptyState({ icon, title, text, action }: { icon: React.ReactNode; title: string; text: string; action?: React.ReactNode }) { return <div className="empty-state"><div className="empty-icon">{icon}</div><h3>{title}</h3><p>{text}</p>{action}</div>; }
function EntityModal({ type, departments, onClose, onSave }: { type: 'department' | 'area' | 'position' | 'person'; departments: Department[]; onClose: () => void; onSave: (value: string, type: 'department' | 'area' | 'position' | 'person', parentId?: string) => void }) { const [value, setValue] = useState(''); const [parent, setParent] = useState(departments[0]?.id || ''); const [area, setArea] = useState(departments[0]?.areas[0]?.id || ''); const label = type === 'department' ? 'departamento' : type === 'area' ? 'área' : type === 'position' ? 'cargo' : 'persona'; const current = departments.find(d => d.id === parent); return <div className="modal-backdrop" onMouseDown={onClose}><div className="modal" onMouseDown={e => e.stopPropagation()}><div className="modal-heading"><div><div className="eyebrow">Nuevo registro</div><h2>Crear {label}</h2></div><button className="icon-button" onClick={onClose}><X size={18} /></button></div><label>Nombre<input autoFocus value={value} onChange={e => setValue(e.target.value)} placeholder={`Nombre del ${label}`} onKeyDown={e => e.key === 'Enter' && onSave(value, type, type === 'department' || type === 'person' ? undefined : type === 'area' ? parent : area)} /></label>{type === 'area' && <label>Departamento<select value={parent} onChange={e => setParent(e.target.value)}>{departments.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}</select></label>}{type === 'position' && <><label>Departamento<select value={parent} onChange={e => { setParent(e.target.value); setArea(departments.find(d => d.id === e.target.value)?.areas[0]?.id || '') }}>{departments.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}</select></label><label>Área opcional<select value={area} onChange={e => setArea(e.target.value)}><option value="">Sin área</option>{current?.areas.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select></label></>}<div className="modal-actions"><button className="secondary-button" onClick={onClose}>Cancelar</button><button className="primary-button" disabled={!value.trim() || (type === 'area' && !parent) || (type === 'position' && !area)} onClick={() => onSave(value, type, type === 'department' || type === 'person' ? undefined : type === 'area' ? parent : area)}>Guardar</button></div></div></div>; }

export default App;

createRoot(document.getElementById('root')!).render(<App />);
