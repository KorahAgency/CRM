import React, { useState, useEffect, useRef } from "react"
import Papa from "papaparse"
import readXlsxFile from "read-excel-file"
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from "recharts"
import { hasSupabase, supabase } from "./src/supabaseClient"

const PINK = '#E8186D', PINK_D = '#B5104F', DARK = '#2C2C2C'
const STAGES = ['Novo Lead','Abordado','Qualificado','Proposta Enviada','Negociação','Convertido','Perdido','Inativo']
const S_BG  = ['#fce4ef','#e3f0fb','#f3e5f8','#fdf5e6','#fdeded','#e8f5ee','#f4f4f4','#fafafa']
const S_CLR = ['#B5104F','#185FA5','#6A1B9A','#D4880A','#A32D2D','#1A6B3C','#555','#888']
const S_CHR = ['#E8186D','#2980B9','#7B3F6E','#D4880A','#C0392B','#1A6B3C','#7F8C8D','#AAAAAA']
const NICHOS = ['Salgados/Buffet','Restaurante','Pizzaria','E-commerce','Odontologia','Estética','Salão de Beleza','Barbearia','Academia','Clínica Médica','Veterinária/Pet Shop','Moda','Mercado/Mercearia','Farmácia','Educação/Cursos','Advocacia','Contabilidade','Imobiliária','Construção/Reforma','Automotivo','Hotelaria/Turismo','Eventos/Festas','Serviços Locais','Infoproduto','Outro']
const SERVICOS = ['Tráfego Pago','Social Media','Site','Design Gráfico','Cardápio Digital','E-commerce','Pacote Completo','Consultoria']

const NICHO_RULES = [
  ['Odontologia',['dentista','odont','ortodont','clinica dental','implante']],
  ['Restaurante',['restaurante','bistro','trattoria','churrascaria','lanchonete','hamburguer','comida','bar e restaurante']],
  ['Pizzaria',['pizzaria','pizza','esfiha']],
  ['Salgados/Buffet',['salgado','salgadinho','coxinha','buffet','festa','confeitaria','doceria','bolo','brigadeiro','crepeiro']],
  ['E-commerce',['ecommerce','e-commerce','loja online','shop','store','marketplace']],
  ['Estética',['estetica','estética','spa','laser','depilacao','depilação','harmonizacao','harmonização','sobrancelha','cilios','cílios']],
  ['Salão de Beleza',['salao','salão','beleza','cabelo','manicure','nail','unha']],
  ['Barbearia',['barbearia','barber','barba']],
  ['Academia',['academia','fitness','crossfit','pilates','personal']],
  ['Clínica Médica',['clinica','clínica','medico','médico','saude','saúde','fisioterapia','psicologia']],
  ['Veterinária/Pet Shop',['pet','veterinaria','veterinária','banho e tosa','animal']],
  ['Moda',['moda','roupa','boutique','calcado','calçado','joia','semijoia']],
  ['Mercado/Mercearia',['mercado','mercearia','hortifruti','padaria','acougue','açougue']],
  ['Farmácia',['farmacia','farmácia','drogaria']],
  ['Educação/Cursos',['curso','escola','colegio','colégio','educacao','educação','idioma']],
  ['Advocacia',['advogado','advocacia','juridico','jurídico']],
  ['Contabilidade',['contabilidade','contador','contabil','contábil']],
  ['Imobiliária',['imobiliaria','imobiliária','imovel','imóvel','corretor']],
  ['Construção/Reforma',['construcao','construção','reforma','arquitetura','engenharia','material de construcao']],
  ['Automotivo',['auto','oficina','mecanica','mecânica','lava jato','veiculo','veículo','carro']],
  ['Hotelaria/Turismo',['hotel','pousada','turismo','viagem','hostel']],
  ['Eventos/Festas',['evento','cerimonial','decoracao','decoração','fotografia','dj']]
]

function normText(v=''){
  return String(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim()
}
function digits(v=''){ return String(v||'').replace(/\D/g,'') }
function phoneDigits(lead={}){
  const d=digits(lead.tel||lead.telefone||lead.phone||lead.wa||lead.whatsapp)
  if(!d)return ''
  if(d.startsWith('55'))return d
  return d.length>=10 ? `55${d}` : d
}
function whatsappLink(lead={}){
  const d=phoneDigits(lead)
  return d&&d.length>=12 ? `https://wa.me/${d}` : ''
}
function classifyNicho(lead={}){
  const current=lead.nicho||lead.categoria||lead.category
  const official=NICHOS.find(n=>normText(n)===normText(current))
  if(official&&official!=='Outro')return official
  const hay=normText([current,lead.nome,lead.end,lead.site,lead.origem,lead.obs].filter(Boolean).join(' '))
  const found=NICHO_RULES.find(([,keys])=>keys.some(k=>hay.includes(normText(k))))
  return found ? found[0] : 'Outro'
}
function leadKey(lead={}){
  const phone=phoneDigits(lead)
  if(phone&&phone.length>=10)return `tel:${phone}`
  const name=normText(lead.nome).replace(/[^a-z0-9]/g,'')
  const site=normText(lead.site).replace(/^https?:\/\//,'').replace(/^www\./,'').split('/')[0]
  return site ? `site:${site}` : `nome:${name}`
}
function enrichLead(lead={}){
  const next={...lead}
  next.nome=String(next.nome||'').trim()
  next.tel=String(next.tel||'').trim()
  next.site=String(next.site||'').trim()
  next.end=String(next.end||'').trim()
  next.nicho=classifyNicho(next)
  next.wa=next.wa&&String(next.wa).startsWith('http') ? next.wa : whatsappLink(next)
  return next
}
function uniqueAgainst(rows,existing=[]){
  const seen=new Set(existing.map(leadKey))
  const unique=[]
  let skipped=0
  rows.forEach(row=>{
    const key=leadKey(row)
    if(seen.has(key)){ skipped++; return }
    seen.add(key); unique.push(row)
  })
  return {unique,skipped}
}
function sheetRowsToObjects(rows=[]){
  const [header=[],...body]=rows
  return body.map(cols=>header.reduce((obj,key,i)=>({...obj,[key||`coluna_${i+1}`]:cols[i]||''}),{}))
}
async function storageGet(key){
  if(typeof window!=='undefined'&&window.storage?.get)return window.storage.get(key,true)
  if(hasSupabase&&supabase){
    const {data,error}=await supabase.from('korah_crm_state').select('value').eq('id',key).maybeSingle()
    if(!error)return {value:data?JSON.stringify(data.value):null,source:'supabase'}
  }
  const value=typeof localStorage!=='undefined' ? localStorage.getItem(key) : null
  return value ? {value,source:'local'} : null
}
async function storageSet(key,value){
  if(typeof window!=='undefined'&&window.storage?.set)return window.storage.set(key,value,true)
  if(hasSupabase&&supabase){
    const parsed=JSON.parse(value)
    const {error}=await supabase.from('korah_crm_state').upsert({id:key,value:parsed,updated_at:new Date().toISOString()})
    if(!error)return
  }
  if(typeof localStorage!=='undefined')localStorage.setItem(key,value)
}

const D_PIPE = [
  {id:1,nome:'Megavip Buffet',nicho:'Salgados/Buffet',tel:'+55 21 97160-1122',wa:'https://wa.me/5521971601122',stage:4,valor:3500,score:95,obs:'Proposta enviada. Decisor: Marcos.',data:'08/07/2026',servicos:'Pacote Completo',site:'megavip.net.br',end:'Rua Pompeu Loureiro, 56 - Copacabana',origem:'Scraping Copacabana'},
  {id:2,nome:'Antonia Salgadinhos',nicho:'Salgados/Buffet',tel:'+55 21 98308-5544',wa:'https://wa.me/5521983085544',stage:2,valor:2500,score:90,obs:'Muito interessada. Só tem Instagram.',data:'08/07/2026',servicos:'Site, Tráfego Pago',site:'',end:'R. Saint Roman, 122 - Copacabana',origem:'Scraping Copacabana'},
  {id:3,nome:'Churrascaria Palace',nicho:'Restaurante',tel:'+55 21 2541-5898',wa:'',stage:5,valor:2200,score:100,obs:'✅ Fechou! R$1.800/mês',data:'01/07/2026',servicos:'Tráfego Pago, Social Media',site:'churrascariapalace.com.br',end:'R. Rodolfo Dantas, 16 - Copacabana',origem:'Scraping Copacabana'},
  {id:4,nome:'Carol Coxinhas',nicho:'Salgados/Buffet',tel:'+55 21 99503-5116',wa:'https://wa.me/5521995035116',stage:1,valor:1500,score:78,obs:'Respondeu positivo. Pediu proposta.',data:'06/07/2026',servicos:'Tráfego Pago',site:'',end:'R. Figueiredo de Magalhães, 581 - Copacabana',origem:'Scraping Copacabana'},
  {id:5,nome:'La Trattoria',nicho:'Restaurante',tel:'+55 21 2255-3319',wa:'',stage:3,valor:1200,score:82,obs:'Proposta enviada. Aguardando.',data:'05/07/2026',servicos:'Tráfego Pago',site:'latrattoriario.com.br',end:'R. Fernando Mendes, 7 - Copacabana',origem:'Scraping Copacabana'},
]
const D_IMP = [
  {id:101,nome:'Salgadinhos do Bilac',nicho:'Salgados/Buffet',tel:'+55 21 97688-1458',wa:'https://wa.me/5521976881458',end:'R. Pedro Lessa, 459 - Duque de Caxias',site:'',origem:'Scraping Duque de Caxias 6km'},
  {id:102,nome:"Tchuco's Salgados",nicho:'Salgados/Buffet',tel:'+55 21 99452-9820',wa:'https://wa.me/5521994529820',end:'R. Saldanha Marinho, 969 - Duque de Caxias',site:'tchucos-salgados.ola.click',origem:'Scraping Duque de Caxias 6km'},
  {id:103,nome:'Confeitaria Coruja',nicho:'Salgados/Buffet',tel:'+55 21 99750-9069',wa:'https://wa.me/5521997509069',end:'Av. Gov. Leonel Brizola, 1790 - Duque de Caxias',site:'',origem:'Scraping Duque de Caxias 6km'},
  {id:104,nome:'Marisa Salgados',nicho:'Salgados/Buffet',tel:'+55 21 99117-4890',wa:'https://wa.me/5521991174890',end:'R. Poços de Caldas, 791 - Jardim Gramacho',site:'instagram.com/marisasalgados1',origem:'Scraping Duque de Caxias 6km'},
  {id:105,nome:'Duda Salgadinhos',nicho:'Salgados/Buffet',tel:'+55 21 99327-9849',wa:'https://wa.me/5521993279849',end:'R. Pedro Lessa, 15 - Jardim Leal, Duque de Caxias',site:'',origem:'Scraping Duque de Caxias 6km'},
  {id:106,nome:'Crepeiros Buffet',nicho:'Salgados/Buffet',tel:'+55 21 99721-6622',wa:'https://wa.me/5521997216622',end:'R. Alm. Gonçalves, 35 - Copacabana',site:'',origem:'Scraping Copacabana 5km'},
  {id:107,nome:'N&F Salgados',nicho:'Salgados/Buffet',tel:'+55 21 98836-8514',wa:'https://wa.me/5521988368514',end:'R. Benfica, 24 - Vila Leopoldina, Duque de Caxias',site:'',origem:'Scraping Duque de Caxias 6km'},
  {id:108,nome:'Humm! Salgados Duque de Caxias',nicho:'Salgados/Buffet',tel:'+55 21 98348-0719',wa:'https://wa.me/5521983480719',end:'R. Barbosa de Araújo, 10 - Parque Felicidade',site:'',origem:'Scraping Duque de Caxias 6km'},
]

const inp = {fontSize:'12px',padding:'6px 10px',border:'0.5px solid var(--border)',borderRadius:'6px',width:'100%',background:'var(--surface-2)',color:'var(--text-primary)',fontFamily:'inherit',boxSizing:'border-box'}
const lbl = {fontSize:'10px',color:'var(--text-muted)',textTransform:'uppercase',letterSpacing:'.4px',marginBottom:'3px',display:'block'}
const TH = {background:DARK,color:'#fff',padding:'8px 12px',textAlign:'left',fontWeight:500,fontSize:'11px',whiteSpace:'nowrap'}
const TD = (alt) => ({padding:'8px 12px',borderBottom:'0.5px solid #f0d0e0',verticalAlign:'middle',background:alt?'#fff5f9':'var(--surface-2)',fontSize:'12px'})
const BTN = (v='pk') => ({background:v==='pk'?PINK:v==='dk'?DARK:v==='gr'?'#25D366':'transparent',color:v==='ghost'?PINK:'#fff',border:v==='ghost'?`1px solid ${PINK}`:'none',padding:'5px 12px',borderRadius:'6px',cursor:'pointer',fontSize:'11px',fontWeight:600,transition:'all .15s'})
const SECBAR = {background:'#fff5f9',borderBottom:'1px solid #f0d0e0',padding:'10px 16px',display:'flex',alignItems:'center',justifyContent:'space-between'}

export default function KorahCRM() {
  const [tab, setTab] = useState('dashboard')
  const [pipeline, _setPipe] = useState([])
  const [imported, _setImp] = useState([])
  const [loading, setLoading] = useState(true)
  const [toast, setToast] = useState(null)
  const [selLead, setSelLead] = useState(null)
  const [pipeView, setPipeView] = useState('table')
  const [dragOver, setDragOver] = useState(false)
  const [preview, setPreview] = useState(null)
  const [filter, setFilter] = useState({q:'',stage:'',nicho:''})
  const [storeMode, setStoreMode] = useState(hasSupabase?'CONECTANDO':'LOCAL')
  const fileRef = useRef()

  useEffect(()=>{ load() },[])

  async function load() {
    try {
      const [pr,ir] = await Promise.allSettled([
        storageGet('korah-crm-pipe'),
        storageGet('korah-crm-imp'),
      ])
      const pipeValue=pr.status==='fulfilled'&&pr.value?.value ? JSON.parse(pr.value.value) : D_PIPE
      const impValue=ir.status==='fulfilled'&&ir.value?.value ? JSON.parse(ir.value.value) : D_IMP
      const connected=(pr.status==='fulfilled'&&pr.value?.source==='supabase')||(ir.status==='fulfilled'&&ir.value?.source==='supabase')
      setStoreMode(connected?'SUPABASE':'LOCAL')
      _setPipe(pipeValue.map(enrichLead))
      _setImp(impValue.map(enrichLead))
    } catch { setStoreMode('LOCAL'); _setPipe(D_PIPE.map(enrichLead)); _setImp(D_IMP.map(enrichLead)) }
    setLoading(false)
  }

  async function savePipe(d){ _setPipe(d); try{ await storageSet('korah-crm-pipe',JSON.stringify(d)) }catch{} }
  async function saveImp(d){  _setImp(d);  try{ await storageSet('korah-crm-imp', JSON.stringify(d)) }catch{} }

  function toast$(msg){ setToast(msg); setTimeout(()=>setToast(null),3000) }

  function mover(id){
    const l=imported.find(x=>x.id===id); if(!l)return
    const lead=enrichLead({id:Date.now(),nome:l.nome,nicho:l.nicho||'Outro',tel:l.tel,wa:l.wa,stage:0,valor:1200,score:70,obs:'',servicos:'',data:new Date().toLocaleDateString('pt-BR'),end:l.end,site:l.site,origem:l.origem})
    if(pipeline.some(p=>leadKey(p)===leadKey(lead))){
      saveImp(imported.filter(x=>x.id!==id))
      toast$(`"${l.nome}" já existe no Pipeline. Duplicado ignorado.`)
      return
    }
    savePipe([lead,...pipeline])
    saveImp(imported.filter(x=>x.id!==id))
    toast$(`"${l.nome}" → Pipeline como Novo Lead!`)
  }
  function moverTodos(){
    const mapped=imported.map(l=>enrichLead({id:Date.now()+Math.random(),nome:l.nome,nicho:l.nicho||'Outro',tel:l.tel,wa:l.wa,stage:0,valor:1200,score:70,obs:'',servicos:'',data:new Date().toLocaleDateString('pt-BR'),end:l.end,site:l.site,origem:l.origem}))
    const {unique:novos,skipped}=uniqueAgainst(mapped,pipeline)
    savePipe([...novos,...pipeline]); saveImp([])
    toast$(`${novos.length} leads movidos. ${skipped} duplicados ignorados.`)
  }
  function chStage(id,s){ savePipe(pipeline.map(l=>l.id===id?{...l,stage:parseInt(s)}:l)) }
  function updLead(id,d){ const lead=enrichLead(d); savePipe(pipeline.map(l=>l.id===id?{...l,...lead}:l)); setSelLead(p=>({...p,...lead})); toast$('Lead atualizado!') }
  function delPipe(id){ savePipe(pipeline.filter(l=>l.id!==id)); setSelLead(null); toast$('Lead removido') }
  function delImp(id){ saveImp(imported.filter(l=>l.id!==id)) }
  function addManual(){
    const n={id:Date.now(),nome:'Novo Lead',nicho:'Outro',tel:'',wa:'',stage:0,valor:1200,score:50,obs:'',servicos:'',data:new Date().toLocaleDateString('pt-BR'),end:'',site:'',origem:'Manual'}
    savePipe([n,...pipeline]); setSelLead(n); setTab('pipeline'); toast$('Lead criado! Clique pra editar.')
  }

  function handleDrop(e){ e.preventDefault(); setDragOver(false); const f=e.dataTransfer.files[0]; if(f)parseFile(f) }
  function parseFile(file){
    const ext=file.name.split('.').pop().toLowerCase()
    if(ext==='csv'){ Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>processImport(r.data,file.name)}) }
    else if(['xlsx','xls'].includes(ext)){
      readXlsxFile(file).then(rows=>processImport(sheetRowsToObjects(rows),file.name)).catch(()=>toast$('Não consegui ler essa planilha. Tente salvar como .xlsx ou CSV.'))
    } else toast$('Use CSV ou Excel (.xlsx)')
  }
  function processImport(rows,filename){
    const get=(row,als)=>{ for(const a of als){ const k=Object.keys(row).find(k=>k.toLowerCase().replace(/[^a-z]/g,'').includes(a)); if(k&&row[k])return String(row[k]) } return '' }
    const mapped=rows.map((row,i)=>enrichLead({id:Date.now()+i,nome:get(row,['nome','name','empresa','title']),tel:get(row,['telefone','phone','tel','celular']),wa:get(row,['whatsapp','wa']),end:get(row,['endereco','address','end','local']),site:get(row,['site','website','url','instagram']),nicho:get(row,['nicho','categoria','category','segmento']),origem:`Importado: ${filename}`})).filter(r=>r.nome)
    const {unique,skipped}=uniqueAgainst(mapped,[...pipeline,...imported])
    setPreview({rows:unique,filename,skipped})
  }
  function confirmImport(){ saveImp([...preview.rows,...imported]); const skipped=preview.skipped||0; setPreview(null); toast$(`${preview.rows.length} leads importados. ${skipped} duplicados ignorados.`) }

  function exportCSV(){
    const csv=Papa.unparse(pipeline.map(l=>({Nome:l.nome,Nicho:l.nicho,Estágio:STAGES[l.stage],Telefone:l.tel,WhatsApp:l.wa||whatsappLink(l),'Valor (R$)':l.valor,Score:l.score,Serviços:l.servicos,Observações:l.obs,'Data Entrada':l.data})))
    const a=Object.assign(document.createElement('a'),{href:URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'})),download:'korah_pipeline.csv'})
    a.click(); toast$('Pipeline exportado!')
  }

  if(loading) return <div style={{padding:'60px',textAlign:'center',color:PINK,fontSize:'14px',fontFamily:'Segoe UI,sans-serif'}}>Carregando CRM Korah Agency...</div>

  const filtered = pipeline.filter(l=>(!filter.q||l.nome.toLowerCase().includes(filter.q.toLowerCase())||l.nicho.toLowerCase().includes(filter.q.toLowerCase())||String(l.tel||'').includes(filter.q))&&(!filter.stage||STAGES[l.stage]===filter.stage)&&(!filter.nicho||l.nicho===filter.nicho))
  const hasMod = selLead||preview

  return (
    <div style={{fontFamily:'Segoe UI,sans-serif',borderRadius:'12px',overflow:'hidden',border:'0.5px solid var(--border)',position:'relative',minHeight: hasMod?'720px':'auto'}}>

      {/* ── Overlay modals (absolute, not fixed) ── */}
      {selLead && (
        <div style={{position:'absolute',inset:0,background:'rgba(0,0,0,.45)',zIndex:100,display:'flex',alignItems:'flex-start',justifyContent:'flex-end'}} onClick={()=>setSelLead(null)}>
          <LeadModal lead={selLead} onSave={d=>updLead(selLead.id,d)} onDelete={()=>delPipe(selLead.id)} onClose={()=>setSelLead(null)} />
        </div>
      )}
      {preview && (
        <div style={{position:'absolute',inset:0,background:'rgba(0,0,0,.45)',zIndex:101,display:'flex',alignItems:'center',justifyContent:'center',padding:'16px'}} onClick={()=>setPreview(null)}>
          <div style={{background:'var(--surface-2)',borderRadius:'12px',padding:'20px',width:'100%',maxWidth:'580px',maxHeight:'70vh',overflowY:'auto'}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:'14px',color:'var(--text-primary)',marginBottom:'2px'}}>Confirmar importação</div>
            <div style={{fontSize:'11px',color:'var(--text-muted)',marginBottom:'14px'}}>{preview.filename} · {preview.rows.length} novos leads · {preview.skipped||0} duplicados ignorados</div>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px',marginBottom:'14px'}}>
              <thead><tr>{['Nome','Telefone','WA?','Endereço'].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
              <tbody>{preview.rows.slice(0,8).map((r,i)=>(
                <tr key={i}>
                  <td style={TD(i%2===0)}>{r.nome}</td>
                  <td style={TD(i%2===0)}>{r.tel}</td>
                  <td style={TD(i%2===0)}>{r.wa?'✅':'—'}</td>
                  <td style={{...TD(i%2===0),maxWidth:'180px',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.end}</td>
                </tr>
              ))}</tbody>
            </table>
            {!preview.rows.length&&<div style={{fontSize:'12px',color:'var(--text-muted)',padding:'18px',textAlign:'center'}}>Todos os leads desse arquivo já estão no CRM.</div>}
            {preview.rows.length>8&&<div style={{fontSize:'10px',color:'var(--text-muted)',marginBottom:'12px'}}>... e mais {preview.rows.length-8} leads</div>}
            <div style={{display:'flex',gap:'8px',justifyContent:'flex-end'}}>
              <button style={{background:'none',border:'1px solid var(--border-strong)',color:'var(--text-primary)',padding:'7px 16px',borderRadius:'8px',cursor:'pointer',fontSize:'12px'}} onClick={()=>setPreview(null)}>Cancelar</button>
              <button disabled={!preview.rows.length} style={{...BTN('pk'),padding:'7px 20px',fontSize:'12px',opacity:preview.rows.length?1:.45,cursor:preview.rows.length?'pointer':'not-allowed'}} onClick={confirmImport}>Importar {preview.rows.length} leads</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Toast ── */}
      {toast && <div style={{position:'absolute',top:'10px',right:'10px',zIndex:200,background:DARK,color:'#fff',padding:'9px 16px',borderRadius:'8px',fontSize:'12px',fontWeight:500,borderLeft:`3px solid ${PINK}`,boxShadow:'0 4px 12px rgba(0,0,0,.15)'}}>{toast}</div>}

      {/* ── Header ── */}
      <div style={{background:PINK,padding:'12px 18px',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
        <div style={{display:'flex',alignItems:'center',gap:'10px'}}>
          <span style={{color:'#fff',fontWeight:700,fontSize:'16px',letterSpacing:'.5px'}}>KORAH AGENCY</span>
          <span style={{background:'rgba(255,255,255,.22)',color:'#fff',fontSize:'10px',padding:'2px 10px',borderRadius:'20px',fontWeight:500}}>CRM PRO</span>
          <span style={{background:'#fff',color:PINK,fontSize:'10px',padding:'2px 10px',borderRadius:'20px',fontWeight:700}}>PÚBLICO</span>
          <span style={{background:storeMode==='SUPABASE'?'#1A6B3C':storeMode==='CONECTANDO'?'#D4880A':'rgba(255,255,255,.22)',color:'#fff',fontSize:'10px',padding:'2px 10px',borderRadius:'20px',fontWeight:700}}>{storeMode}</span>
        </div>
        <div style={{display:'flex',gap:'8px',alignItems:'center'}}>
          <span style={{color:'rgba(255,255,255,.8)',fontSize:'11px'}}>{pipeline.length} leads · {imported.length} p/ abordar</span>
          <button style={{background:'rgba(255,255,255,.18)',border:'1px solid rgba(255,255,255,.3)',color:'#fff',padding:'4px 12px',borderRadius:'6px',cursor:'pointer',fontSize:'11px',fontWeight:500}} onClick={addManual}>+ Lead</button>
          <button style={{background:'rgba(255,255,255,.18)',border:'1px solid rgba(255,255,255,.3)',color:'#fff',padding:'4px 12px',borderRadius:'6px',cursor:'pointer',fontSize:'11px',fontWeight:500}} onClick={exportCSV}>Exportar CSV</button>
        </div>
      </div>

      {/* ── Tabs ── */}
      <div style={{background:DARK,display:'flex',overflowX:'auto'}}>
        {[['dashboard','📊 Dashboard'],['pipeline','📋 Pipeline'],['import','📥 Importar CSV/Excel'],['follow','📅 Follow-Up'],['scripts','💬 Scripts'],['hist','📝 Histórico']].map(([k,l])=>(
          <button key={k} onClick={()=>setTab(k)} style={{padding:'10px 14px',fontSize:'12px',color:tab===k?'#fff':'rgba(255,255,255,.5)',background:'none',border:'none',borderBottom:tab===k?`2px solid ${PINK}`:'2px solid transparent',cursor:'pointer',fontWeight:500,whiteSpace:'nowrap',transition:'all .15s'}}>
            {l}
          </button>
        ))}
      </div>

      {/* ── Content ── */}
      <div>
        {tab==='dashboard' && <TabDash pipeline={pipeline} imported={imported} setSelLead={setSelLead} setTab={setTab} />}
        {tab==='pipeline' && <TabPipe pipeline={filtered} full={pipeline} filter={filter} setFilter={setFilter} pipeView={pipeView} setPipeView={setPipeView} setSelLead={setSelLead} chStage={chStage} delPipe={delPipe} />}
        {tab==='import' && <TabImport imported={imported} mover={mover} moverTodos={moverTodos} delImp={delImp} dragOver={dragOver} setDragOver={setDragOver} handleDrop={handleDrop} parseFile={parseFile} fileRef={fileRef} />}
        {tab==='follow' && <TabFollow pipeline={pipeline} />}
        {tab==='scripts' && <TabScripts />}
        {tab==='hist' && <TabHist pipeline={pipeline} />}
      </div>
    </div>
  )
}

// ─── Dashboard ──────────────────────────────────────────────────
function TabDash({pipeline,imported,setSelLead,setTab}){
  const total=pipeline.length
  const novos=pipeline.filter(p=>p.stage===0).length
  const and=pipeline.filter(p=>p.stage>=1&&p.stage<=4).length
  const conv=pipeline.filter(p=>p.stage===5)
  const fat=conv.reduce((a,b)=>a+(b.valor||0),0)
  const chartData=STAGES.slice(0,6).map((s,i)=>({name:s.length>9?s.slice(0,8)+'…':s,full:s,value:pipeline.filter(p=>p.stage===i).length}))
  const top=[...pipeline].filter(p=>p.stage<5&&p.stage>0).sort((a,b)=>(b.score||0)-(a.score||0)).slice(0,5)
  const semSite=pipeline.filter(p=>!p.site&&p.stage<5).length
  const porNicho=NICHOS.map(n=>({n,total:pipeline.filter(p=>p.nicho===n).length,abertos:pipeline.filter(p=>p.nicho===n&&p.stage<5).length})).filter(x=>x.total).sort((a,b)=>b.total-a.total).slice(0,8)
  return (
    <div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'10px',padding:'14px 14px 0'}}>
        {[
          {l:'Total no pipeline',v:total,s:`${imported.length} aguardando abordagem`,c:DARK},
          {l:'Novos leads',v:novos,s:'prontos para abordar',c:PINK},
          {l:'Em andamento',v:and,s:`${pipeline.filter(p=>p.stage===6).length} perdidos`,c:'#7B3F6E'},
          {l:'Convertidos',v:conv.length,s:`R$ ${fat.toLocaleString('pt-BR')}/mês`,c:'#1A6B3C'},
        ].map(({l,v,s,c},i)=>(
          <div key={i} style={{background:'var(--surface-1)',borderRadius:'10px',padding:'14px',borderLeft:`3px solid ${c}`}}>
            <div style={{fontSize:'10px',color:'var(--text-muted)',textTransform:'uppercase',letterSpacing:'.4px',marginBottom:'4px'}}>{l}</div>
            <div style={{fontSize:'28px',fontWeight:700,color:c,lineHeight:1}}>{v}</div>
            <div style={{fontSize:'10px',color:'var(--text-muted)',marginTop:'4px'}}>{s}</div>
          </div>
        ))}
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'10px',padding:'10px 14px 14px'}}>
        <div style={{background:'var(--surface-1)',borderRadius:'10px',padding:'14px',border:'0.5px solid var(--border)'}}>
          <div style={{fontSize:'12px',fontWeight:600,color:DARK,marginBottom:'10px'}}>Leads por estágio</div>
          <ResponsiveContainer width="100%" height={170}>
            <BarChart data={chartData} margin={{top:0,right:0,left:-24,bottom:0}}>
              <XAxis dataKey="name" tick={{fontSize:9}} />
              <YAxis tick={{fontSize:9}} />
              <Tooltip formatter={(v,_,p)=>[v,p.payload.full]} />
              <Bar dataKey="value" radius={[4,4,0,0]}>{chartData.map((_,i)=><Cell key={i} fill={S_CHR[i]}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={{background:'var(--surface-1)',borderRadius:'10px',padding:'14px',border:'0.5px solid var(--border)'}}>
          <div style={{fontSize:'12px',fontWeight:600,color:DARK,marginBottom:'10px'}}>Top leads por score</div>
          {!top.length&&<div style={{fontSize:'11px',color:'var(--text-muted)'}}>Sem leads em andamento.</div>}
          {top.map(l=>(
            <div key={l.id} onClick={()=>{setSelLead(l);setTab('pipeline')}} style={{display:'flex',alignItems:'center',gap:'10px',padding:'7px 0',borderBottom:'0.5px solid var(--border)',cursor:'pointer'}}>
              <div style={{width:30,height:30,borderRadius:'50%',background:S_BG[l.stage],display:'flex',alignItems:'center',justifyContent:'center',fontSize:'11px',fontWeight:700,color:S_CLR[l.stage],flexShrink:0}}>{l.score}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:'12px',fontWeight:600,color:'var(--text-primary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.nome}</div>
                <div style={{fontSize:'10px',color:'var(--text-muted)'}}>{STAGES[l.stage]}</div>
              </div>
              <div style={{fontSize:'11px',fontWeight:700,color:PINK,flexShrink:0}}>R$ {(l.valor||0).toLocaleString('pt-BR')}</div>
            </div>
          ))}
          {semSite>0&&<div style={{marginTop:'10px',background:'#fff5f9',borderRadius:'8px',padding:'8px 10px',fontSize:'11px',color:PINK_D,fontWeight:500}}>💡 {semSite} leads sem site — oportunidade de venda!</div>}
        </div>
      </div>
      <div style={{padding:'0 14px 14px'}}>
        <div style={{background:'var(--surface-1)',borderRadius:'10px',padding:'14px',border:'0.5px solid var(--border)'}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'10px'}}>
            <div style={{fontSize:'12px',fontWeight:600,color:DARK}}>Visão por nicho</div>
            <button style={BTN('ghost')} onClick={()=>setTab('pipeline')}>Abrir categorias</button>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'8px'}}>
            {porNicho.map(x=>(
              <div key={x.n} style={{border:'0.5px solid var(--border)',borderRadius:'8px',padding:'10px',background:'var(--surface-2)'}}>
                <div style={{fontSize:'11px',fontWeight:700,color:'var(--text-primary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{x.n}</div>
                <div style={{fontSize:'20px',fontWeight:700,color:PINK,lineHeight:1.2,marginTop:'4px'}}>{x.total}</div>
                <div style={{fontSize:'10px',color:'var(--text-muted)'}}>{x.abertos} em aberto</div>
              </div>
            ))}
            {!porNicho.length&&<div style={{fontSize:'11px',color:'var(--text-muted)'}}>Sem leads categorizados ainda.</div>}
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Pipeline ───────────────────────────────────────────────────
function TabPipe({pipeline,full,filter,setFilter,pipeView,setPipeView,setSelLead,chStage,delPipe}){
  return (
    <div>
      <div style={{...SECBAR,gap:'8px',flexWrap:'wrap'}}>
        <span style={{fontSize:'11px',fontWeight:700,color:DARK,textTransform:'uppercase',letterSpacing:'.4px'}}>Pipeline ({pipeline.length}/{full.length})</span>
        <div style={{display:'flex',gap:'6px',alignItems:'center',flex:1,justifyContent:'flex-end',flexWrap:'wrap'}}>
          <input value={filter.q} onChange={e=>setFilter(f=>({...f,q:e.target.value}))} placeholder="Buscar..." style={{...inp,width:'150px',padding:'5px 9px'}} />
          <select value={filter.stage} onChange={e=>setFilter(f=>({...f,stage:e.target.value}))} style={{...inp,width:'auto',padding:'5px 8px'}}>
            <option value="">Todos</option>
            {STAGES.map(s=><option key={s}>{s}</option>)}
          </select>
          <select value={filter.nicho} onChange={e=>setFilter(f=>({...f,nicho:e.target.value}))} style={{...inp,width:'auto',padding:'5px 8px'}}>
            <option value="">Todos os nichos</option>
            {NICHOS.map(n=><option key={n}>{n}</option>)}
          </select>
          <div style={{display:'flex',gap:'3px'}}>
            {[['table','Tabela'],['kanban','Kanban'],['category','Categorias']].map(([k,l])=>(
              <button key={k} onClick={()=>setPipeView(k)} style={{...BTN(pipeView===k?'dk':'ghost'),padding:'4px 10px'}}>{l}</button>
            ))}
          </div>
        </div>
      </div>
      {pipeView==='table' ? (
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>{['Empresa / Nicho','Estágio','WhatsApp','Valor','Score','Obs',''].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>
              {!pipeline.length&&<tr><td colSpan={7} style={{padding:'24px',textAlign:'center',color:'var(--text-muted)',fontSize:'12px'}}>Nenhum lead encontrado.</td></tr>}
              {pipeline.map((l,i)=>(
                <tr key={l.id} style={{cursor:'pointer'}} onClick={()=>setSelLead(l)}>
                  <td style={TD(i%2===0)}>
                    <div style={{fontWeight:600,fontSize:'12px',color:'var(--text-primary)'}}>{l.nome}</div>
                    <div style={{fontSize:'10px',color:'var(--text-muted)'}}>{l.nicho} · {l.tel}</div>
                  </td>
                  <td style={TD(i%2===0)} onClick={e=>e.stopPropagation()}>
                    <select value={l.stage} onChange={e=>chStage(l.id,e.target.value)} style={{fontSize:'11px',border:'none',borderRadius:'6px',padding:'3px 6px',background:S_BG[l.stage],color:S_CLR[l.stage],fontWeight:700,cursor:'pointer'}}>
                      {STAGES.map((s,i)=><option key={i} value={i}>{s}</option>)}
                    </select>
                  </td>
                  <td style={TD(i%2===0)} onClick={e=>e.stopPropagation()}>
                    {(l.wa||whatsappLink(l))?<a href={l.wa||whatsappLink(l)} target="_blank" rel="noreferrer" style={{color:PINK,fontSize:'11px',textDecoration:'none',fontWeight:500}}>Abrir WA →</a>:<span style={{color:'var(--text-muted)',fontSize:'10px'}}>—</span>}
                  </td>
                  <td style={{...TD(i%2===0),fontWeight:700,color:PINK,fontSize:'12px'}}>R$ {(l.valor||0).toLocaleString('pt-BR')}</td>
                  <td style={TD(i%2===0)}>
                    <ScoreBadge score={l.score||0} />
                  </td>
                  <td style={{...TD(i%2===0),maxWidth:'150px',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:'10px',color:'var(--text-secondary)'}}>{l.obs}</td>
                  <td style={TD(i%2===0)} onClick={e=>e.stopPropagation()}>
                    <button onClick={()=>delPipe(l.id)} style={{background:'none',border:'none',cursor:'pointer',color:'var(--text-muted)',fontSize:'14px',padding:'2px 5px'}}>✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : pipeView==='kanban' ? (
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'8px',padding:'12px'}}>
          {STAGES.map((stage,si)=>{
            const ls=pipeline.filter(l=>l.stage===si)
            return (
              <div key={si} style={{background:'var(--surface-1)',borderRadius:'8px',overflow:'hidden'}}>
                <div style={{padding:'8px 10px',fontSize:'10px',fontWeight:700,background:S_BG[si],color:S_CLR[si],display:'flex',justifyContent:'space-between'}}>
                  <span>{stage}</span><span>{ls.length}</span>
                </div>
                <div style={{padding:'6px',minHeight:'80px'}}>
                  {!ls.length&&<div style={{padding:'10px',fontSize:'10px',color:'var(--text-muted)',textAlign:'center'}}>—</div>}
                  {ls.map(l=>(
                    <div key={l.id} onClick={()=>setSelLead(l)} style={{background:'var(--surface-2)',border:'0.5px solid var(--border)',borderRadius:'6px',padding:'8px',marginBottom:'5px',cursor:'pointer'}}>
                      <div style={{fontSize:'11px',fontWeight:600,color:'var(--text-primary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',marginBottom:'2px'}}>{l.nome}</div>
                      <div style={{fontSize:'10px',color:'var(--text-muted)',marginBottom:'3px'}}>{l.nicho}</div>
                      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                        <span style={{fontSize:'11px',fontWeight:700,color:PINK}}>R$ {(l.valor||0).toLocaleString('pt-BR')}</span>
                        <ScoreBadge score={l.score||0} small />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )
          })}
        </div>
      ) : (
        <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'10px',padding:'12px'}}>
          {NICHOS.map(n=>{
            const ls=pipeline.filter(l=>l.nicho===n)
            if(!ls.length)return null
            return (
              <div key={n} style={{background:'var(--surface-1)',border:'0.5px solid var(--border)',borderRadius:'8px',overflow:'hidden'}}>
                <div style={{padding:'9px 10px',background:'#fff5f9',display:'flex',justifyContent:'space-between',gap:'8px'}}>
                  <span style={{fontSize:'11px',fontWeight:700,color:PINK_D,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{n}</span>
                  <span style={{fontSize:'11px',fontWeight:700,color:DARK}}>{ls.length}</span>
                </div>
                <div style={{padding:'6px'}}>
                  {ls.slice(0,6).map(l=>(
                    <div key={l.id} onClick={()=>setSelLead(l)} style={{display:'flex',alignItems:'center',gap:'8px',padding:'7px 6px',borderBottom:'0.5px solid var(--border)',cursor:'pointer'}}>
                      <ScoreBadge score={l.score||0} small />
                      <div style={{minWidth:0,flex:1}}>
                        <div style={{fontSize:'11px',fontWeight:600,color:'var(--text-primary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.nome}</div>
                        <div style={{fontSize:'10px',color:'var(--text-muted)'}}>{STAGES[l.stage]}</div>
                      </div>
                    </div>
                  ))}
                  {ls.length>6&&<div style={{fontSize:'10px',color:'var(--text-muted)',padding:'6px'}}>+ {ls.length-6} leads nessa categoria</div>}
                </div>
              </div>
            )
          })}
          {!pipeline.length&&<div style={{padding:'24px',textAlign:'center',color:'var(--text-muted)',fontSize:'12px'}}>Nenhum lead encontrado.</div>}
        </div>
      )}
    </div>
  )
}

// ─── Import ─────────────────────────────────────────────────────
function TabImport({imported,mover,moverTodos,delImp,dragOver,setDragOver,handleDrop,parseFile,fileRef}){
  return (
    <div style={{padding:'14px'}}>
      <div onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)} onDrop={handleDrop} onClick={()=>fileRef.current.click()} style={{border:`2px dashed ${dragOver?PINK:'var(--border-strong)'}`,borderRadius:'12px',padding:'28px',textAlign:'center',background:dragOver?'#fce4ef':'var(--surface-1)',cursor:'pointer',marginBottom:'14px',transition:'all .2s'}}>
        <div style={{fontSize:'28px',marginBottom:'8px'}}>📂</div>
        <div style={{fontSize:'14px',fontWeight:600,color:dragOver?PINK:'var(--text-primary)',marginBottom:'3px'}}>Arraste o arquivo aqui</div>
        <div style={{fontSize:'11px',color:'var(--text-muted)',marginBottom:'12px'}}>Suporta CSV e Excel (.xlsx) gerados pelo Claude — colunas: Nome, Telefone, WhatsApp, Endereço, Site</div>
        <button style={{...BTN('pk'),fontSize:'12px',padding:'6px 16px'}} onClick={e=>{e.stopPropagation();fileRef.current.click()}}>Escolher arquivo</button>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{display:'none'}} onChange={e=>{if(e.target.files[0])parseFile(e.target.files[0])}} />
      </div>
      <div style={SECBAR}>
        <span style={{fontSize:'11px',fontWeight:700,color:DARK,textTransform:'uppercase',letterSpacing:'.4px'}}>Aguardando abordagem ({imported.length})</span>
        {imported.length>0&&<button style={BTN('pk')} onClick={moverTodos}>Mover todos → Pipeline</button>}
      </div>
      {!imported.length&&<div style={{padding:'24px',textAlign:'center',color:'var(--text-muted)',fontSize:'12px'}}>Nenhum lead importado. Arraste um arquivo CSV ou Excel acima.</div>}
      {imported.length>0&&(
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>{['Empresa','Nicho','Telefone','WhatsApp','Site','Origem','Ação'].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>{imported.map((l,i)=>(
              <tr key={l.id}>
                <td style={TD(i%2===0)}><div style={{fontWeight:600,fontSize:'12px',color:'var(--text-primary)'}}>{l.nome}</div><div style={{fontSize:'10px',color:'var(--text-muted)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'180px'}}>{l.end}</div></td>
                <td style={TD(i%2===0)}><span style={{background:'#f0e8f5',color:'#6A1B9A',fontSize:'10px',padding:'2px 7px',borderRadius:'10px'}}>{l.nicho}</span></td>
                <td style={{...TD(i%2===0),fontSize:'11px'}}>{l.tel}</td>
                <td style={TD(i%2===0)}>{(l.wa||whatsappLink(l))?<a href={l.wa||whatsappLink(l)} target="_blank" rel="noreferrer" style={{color:PINK,fontSize:'11px',textDecoration:'none',fontWeight:500}}>Abrir WA →</a>:<span style={{color:'var(--text-muted)',fontSize:'10px'}}>—</span>}</td>
                <td style={{...TD(i%2===0),fontSize:'10px',color:'var(--text-muted)'}}>{l.site||'—'}</td>
                <td style={{...TD(i%2===0),fontSize:'10px',color:'var(--text-muted)'}}>{l.origem}</td>
                <td style={TD(i%2===0)}>
                  <div style={{display:'flex',gap:'4px'}}>
                    <button style={{...BTN('pk'),fontSize:'11px',padding:'4px 10px'}} onClick={()=>mover(l.id)}>→ Pipeline</button>
                    <button style={{background:'none',border:'none',cursor:'pointer',color:'var(--text-muted)',fontSize:'13px',padding:'3px 6px'}} onClick={()=>delImp(l.id)}>✕</button>
                  </div>
                </td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  )
}

// ─── Follow-Up ──────────────────────────────────────────────────
function TabFollow({pipeline}){
  const urgentes=pipeline.filter(p=>p.stage===4||p.stage===3).sort((a,b)=>(b.score||0)-(a.score||0))
  const abordados=pipeline.filter(p=>p.stage===1)
  const novos=pipeline.filter(p=>p.stage===0)
  const Section=({title,color,leads,roteiro})=>!leads.length?null:(
    <div style={{marginBottom:'4px'}}>
      <div style={{padding:'8px 16px',fontSize:'11px',fontWeight:700,color,borderLeft:`3px solid ${color}`,background:'var(--surface-1)'}}>{title} ({leads.length})</div>
      <div style={{overflowX:'auto'}}>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Empresa','Estágio','WhatsApp','Score','Roteiro sugerido'].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
          <tbody>{leads.map((l,i)=>(
            <tr key={l.id}>
              <td style={TD(i%2===0)}><div style={{fontWeight:600,fontSize:'12px',color:'var(--text-primary)'}}>{l.nome}</div><div style={{fontSize:'10px',color:'var(--text-muted)'}}>{l.nicho}</div></td>
              <td style={TD(i%2===0)}><span style={{padding:'2px 8px',borderRadius:'10px',fontSize:'10px',fontWeight:700,background:S_BG[l.stage],color:S_CLR[l.stage]}}>{STAGES[l.stage]}</span></td>
              <td style={TD(i%2===0)}>{(l.wa||whatsappLink(l))?<a href={l.wa||whatsappLink(l)} target="_blank" rel="noreferrer" style={{color:PINK,fontSize:'11px',textDecoration:'none',fontWeight:500}}>Abrir WA →</a>:'—'}</td>
              <td style={TD(i%2===0)}><ScoreBadge score={l.score||0} /></td>
              <td style={{...TD(i%2===0),fontSize:'10px',color:'var(--text-secondary)',maxWidth:'220px'}}>{roteiro(l)}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  )
  if(!urgentes.length&&!abordados.length&&!novos.length) return <div style={{padding:'32px',textAlign:'center',color:'var(--text-muted)',fontSize:'13px'}}>Nenhum follow-up pendente! 🎉</div>
  return (
    <div>
      <div style={SECBAR}><span style={{fontSize:'11px',fontWeight:700,color:DARK,textTransform:'uppercase',letterSpacing:'.4px'}}>Follow-up prioritário ({urgentes.length+abordados.length+novos.length})</span></div>
      <Section title="🔴 Proposta / Negociação — contatar hoje" color="#C0392B" leads={urgentes} roteiro={l=>`Oi ${l.nome.split(' ')[0]}! Passando pra ver se teve chance de analisar nossa proposta. Posso te mostrar um case? 🚀`} />
      <Section title="🟡 Abordados — sem resposta" color="#D4880A" leads={abordados} roteiro={l=>`Sei que o dia a dia é corrido 😅 Tenho um case de ${l.nicho} que cresceu 40% em 2 meses com a Korah. Te conto rapidinho?`} />
      <Section title="🟢 Novos leads — abordar agora" color={PINK} leads={novos} roteiro={l=>`Oi! Vi o ${l.nome} no Google. Sou da Korah Agency — trabalho com marketing pra ${l.nicho}. Posso te mostrar como funciona? 😊`} />
    </div>
  )
}

// ─── Scripts ────────────────────────────────────────────────────
function TabScripts(){
  const base=[
    {id:1,t:'Primeiro contato — Salgados/Buffet',c:PINK,txt:`Olá [NOME]! Tudo bem? 😊\n\nVi o [NEGÓCIO] no Google e fiquei impressionado(a) com o trabalho de vocês!\n\nSou da Korah Agency — trabalhamos com marketing digital pra negócios de alimentação. Já ajudamos mais de 80 salgadeiros e buffets a crescer vendas com tráfego pago.\n\nVocê toparia uma conversa rápida de 10 minutos? 🎯`},
    {id:2,t:'Follow-up — sem resposta (3 dias)',c:'#2980B9',txt:`Oi [NOME]! Sei que o dia a dia é corrido 😅\n\nPassei pra retomar minha mensagem. Tenho um case de salgaderia aqui do RJ que aumentou o faturamento em 40% em 2 meses com a gente.\n\nPosso te contar em 5 minutinhos? ☎️`},
    {id:3,t:'Envio de proposta',c:'#D4880A',txt:`Oi [NOME]! Conforme conversamos, segue a proposta da Korah Agency:\n\n📦 Pacote: [NOME DO PACOTE]\n💰 Investimento: R$ [VALOR]/mês\n\n✅ Inclui:\n• Gestão de tráfego pago (Meta Ads)\n• [OUTROS SERVIÇOS]\n• Relatório quinzenal\n• Suporte via WhatsApp\n\nDisponível pra uma call amanhã? 🚀`},
    {id:4,t:'Reengajamento — lead perdido (90 dias)',c:'#7F8C8D',txt:`Oi [NOME], faz um tempão! 👋\n\nConversamos antes sobre marketing pro [NEGÓCIO], mas entendo que não era o momento.\n\nNossos resultados com [NICHO] estão incríveis — uma empresa similar triplicou os pedidos em 3 meses.\n\nUma conversa de 10 minutos? Sem pressão! 😊`},
    {id:5,t:'Pós-fechamento — boas-vindas Korah',c:'#1A6B3C',txt:`Oi [NOME]! Seja muito bem-vindo(a) à família Korah Agency! 🎉\n\nPróximos passos:\n1️⃣ Reunião de onboarding: [DATA]\n2️⃣ Preenchimento do briefing: [LINK]\n3️⃣ Acesso às redes e contas de anúncio\n4️⃣ Início das campanhas: [DATA]\n\nA Korah vai fazer seu negócio crescer! 🚀✨`},
  ]
  const [scripts,setScripts]=useState(base)
  useEffect(()=>{ storageGet('korah-scripts').then(r=>{ if(r?.value)setScripts(JSON.parse(r.value)) }).catch(()=>{}) },[])
  function persist(next){ setScripts(next); storageSet('korah-scripts',JSON.stringify(next)).catch(()=>{}) }
  function upd(id,d){ persist(scripts.map(s=>s.id===id?{...s,...d}:s)) }
  function add(){
    persist([{id:Date.now(),t:'Novo script',c:PINK,txt:'Oi [NOME], tudo bem?\n\nEscreva aqui sua mensagem.'},...scripts])
  }
  function del(id){ persist(scripts.filter(s=>s.id!==id)) }
  function reset(){ persist(base) }
  return (
    <div>
      <div style={SECBAR}>
        <span style={{fontSize:'11px',fontWeight:700,color:DARK,textTransform:'uppercase',letterSpacing:'.4px'}}>Scripts editáveis ({scripts.length})</span>
        <div style={{display:'flex',gap:'8px'}}>
          <button style={BTN('ghost')} onClick={reset}>Restaurar padrão</button>
          <button style={BTN('pk')} onClick={add}>+ Script</button>
        </div>
      </div>
      <div style={{padding:'14px',display:'flex',flexDirection:'column',gap:'10px'}}>
        {scripts.map(s=><ScriptCard key={s.id} script={s} onChange={d=>upd(s.id,d)} onDelete={()=>del(s.id)} />)}
      </div>
    </div>
  )
}

// ─── Histórico ──────────────────────────────────────────────────
function TabHist({pipeline}){
  const [hist,setHist]=useState([])
  const [form,setForm]=useState({empresa:'',canal:'WhatsApp',tipo:'Primeiro contato',resultado:'',obs:''})
  const canais=['WhatsApp','Telefone','E-mail','Instagram DM','Presencial']
  const tipos=['Primeiro contato','Follow-up','Envio proposta','Negociação','Fechamento','Pós-venda']
  const resultados=['Sem resposta','Respondeu - positivo','Respondeu - negativo','Agendou reunião','Pediu proposta','Fechou','Perdeu interesse']
  useEffect(()=>{ storageGet('korah-hist').then(r=>{ if(r?.value)setHist(JSON.parse(r.value)) }).catch(()=>{}) },[])
  function add(){
    if(!form.empresa)return
    const r={id:Date.now(),data:new Date().toLocaleDateString('pt-BR'),...form}
    const h=[r,...hist]
    setHist(h); storageSet('korah-hist',JSON.stringify(h)).catch(()=>{})
    setForm(f=>({...f,empresa:'',resultado:'',obs:''}))
  }
  function del(id){ const h=hist.filter(x=>x.id!==id); setHist(h); storageSet('korah-hist',JSON.stringify(h)).catch(()=>{}) }
  return (
    <div>
      <div style={SECBAR}><span style={{fontSize:'11px',fontWeight:700,color:DARK,textTransform:'uppercase',letterSpacing:'.4px'}}>Registrar interação</span></div>
      <div style={{padding:'12px 14px',background:'var(--surface-1)',borderBottom:'0.5px solid var(--border)',display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr 2fr auto',gap:'8px',alignItems:'end'}}>
        <div><label style={lbl}>Empresa</label>
          <select style={inp} value={form.empresa} onChange={e=>setForm(f=>({...f,empresa:e.target.value}))}>
            <option value="">Selecionar...</option>
            {pipeline.map(l=><option key={l.id}>{l.nome}</option>)}
          </select>
        </div>
        <div><label style={lbl}>Canal</label><select style={inp} value={form.canal} onChange={e=>setForm(f=>({...f,canal:e.target.value}))}>{canais.map(c=><option key={c}>{c}</option>)}</select></div>
        <div><label style={lbl}>Tipo</label><select style={inp} value={form.tipo} onChange={e=>setForm(f=>({...f,tipo:e.target.value}))}>{tipos.map(t=><option key={t}>{t}</option>)}</select></div>
        <div><label style={lbl}>Resultado</label><select style={inp} value={form.resultado} onChange={e=>setForm(f=>({...f,resultado:e.target.value}))}><option value="">Selecionar...</option>{resultados.map(r=><option key={r}>{r}</option>)}</select></div>
        <div><label style={lbl}>Anotação</label><input style={inp} value={form.obs} onChange={e=>setForm(f=>({...f,obs:e.target.value}))} placeholder="O que foi dito..." /></div>
        <button style={{...BTN('pk'),padding:'6px 14px',height:'32px',marginTop:'auto'}} onClick={add}>Registrar</button>
      </div>
      <div style={{overflowX:'auto'}}>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Data','Empresa','Canal','Tipo','Resultado','Anotação',''].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
          <tbody>
            {!hist.length&&<tr><td colSpan={7} style={{padding:'24px',textAlign:'center',color:'var(--text-muted)',fontSize:'12px'}}>Nenhuma interação registrada ainda.</td></tr>}
            {hist.map((h,i)=>(
              <tr key={h.id}>
                <td style={{...TD(i%2===0),whiteSpace:'nowrap',fontSize:'11px'}}>{h.data}</td>
                <td style={{...TD(i%2===0),fontWeight:600,fontSize:'12px',color:'var(--text-primary)'}}>{h.empresa}</td>
                <td style={{...TD(i%2===0),fontSize:'11px'}}>{h.canal}</td>
                <td style={{...TD(i%2===0),fontSize:'11px'}}>{h.tipo}</td>
                <td style={TD(i%2===0)}><span style={{fontSize:'10px',fontWeight:600,color:h.resultado?.includes('positivo')||h.resultado==='Fechou'?'#1A6B3C':h.resultado?.includes('negativo')||h.resultado==='Perdeu interesse'?'#C0392B':'#D4880A'}}>{h.resultado}</span></td>
                <td style={{...TD(i%2===0),fontSize:'11px',color:'var(--text-secondary)',maxWidth:'200px'}}>{h.obs}</td>
                <td style={TD(i%2===0)}><button onClick={()=>del(h.id)} style={{background:'none',border:'none',cursor:'pointer',color:'var(--text-muted)',fontSize:'13px'}}>✕</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ─── Lead Modal (slide-in panel) ─────────────────────────────────
function LeadModal({lead,onSave,onDelete,onClose}){
  const [f,setF]=useState({...lead})
  const u=(k,v)=>setF(p=>{
    const next={...p,[k]:v}
    if(k==='tel'&&(!p.wa||String(p.wa).includes('wa.me/')))next.wa=whatsappLink(next)
    return next
  })
  return (
    <div style={{width:'370px',height:'100%',overflowY:'auto',background:'var(--surface-2)',padding:'18px',boxSizing:'border-box'}} onClick={e=>e.stopPropagation()}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'14px'}}>
        <div><div style={{fontSize:'14px',fontWeight:700,color:'var(--text-primary)'}}>{f.nome}</div><div style={{fontSize:'10px',color:'var(--text-muted)'}}>{f.nicho} · {f.data}</div></div>
        <button onClick={onClose} style={{background:'none',border:'none',cursor:'pointer',fontSize:'18px',color:'var(--text-muted)',padding:'4px'}}>✕</button>
      </div>
      <div style={{background:S_BG[f.stage],borderRadius:'8px',padding:'10px 12px',marginBottom:'12px',display:'flex',alignItems:'center',gap:'6px'}}>
        <span style={{fontSize:'11px',fontWeight:700,color:S_CLR[f.stage]}}>Estágio:</span>
        <select value={f.stage} onChange={e=>u('stage',parseInt(e.target.value))} style={{fontSize:'12px',border:'none',background:'transparent',color:S_CLR[f.stage],fontWeight:700,cursor:'pointer'}}>
          {STAGES.map((s,i)=><option key={i} value={i}>{s}</option>)}
        </select>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px',marginBottom:'8px'}}>
        <div><label style={lbl}>Nome / Empresa</label><input style={inp} value={f.nome} onChange={e=>u('nome',e.target.value)} /></div>
        <div><label style={lbl}>Nicho</label><select style={inp} value={f.nicho} onChange={e=>u('nicho',e.target.value)}>{NICHOS.map(n=><option key={n}>{n}</option>)}</select></div>
        <div><label style={lbl}>Telefone</label><input style={inp} value={f.tel||''} onChange={e=>u('tel',e.target.value)} /></div>
        <div><label style={lbl}>WhatsApp link</label><input style={inp} value={f.wa||whatsappLink(f)} onChange={e=>u('wa',e.target.value)} /></div>
        <div><label style={lbl}>Valor estimado (R$)</label><input style={inp} type="number" value={f.valor||0} onChange={e=>u('valor',parseInt(e.target.value)||0)} /></div>
        <div><label style={lbl}>Score (0–100)</label><input style={inp} type="number" min={0} max={100} value={f.score||0} onChange={e=>u('score',parseInt(e.target.value)||0)} /></div>
      </div>
      <div style={{marginBottom:'8px'}}><label style={lbl}>Serviços de interesse</label><select style={inp} value={f.servicos||''} onChange={e=>u('servicos',e.target.value)}><option value="">Selecionar...</option>{SERVICOS.map(s=><option key={s}>{s}</option>)}</select></div>
      <div style={{marginBottom:'8px'}}><label style={lbl}>Site</label><input style={inp} value={f.site||''} onChange={e=>u('site',e.target.value)} /></div>
      <div style={{marginBottom:'8px'}}><label style={lbl}>Endereço</label><input style={inp} value={f.end||''} onChange={e=>u('end',e.target.value)} /></div>
      <div style={{marginBottom:'12px'}}><label style={lbl}>Observações</label><textarea style={{...inp,height:'70px',resize:'vertical'}} value={f.obs||''} onChange={e=>u('obs',e.target.value)} /></div>
      {(f.wa||whatsappLink(f))&&<a href={f.wa||whatsappLink(f)} target="_blank" rel="noreferrer" style={{display:'block',background:'#25D366',color:'#fff',padding:'8px 14px',borderRadius:'8px',textDecoration:'none',textAlign:'center',fontSize:'12px',fontWeight:600,marginBottom:'10px'}}>Abrir no WhatsApp</a>}
      <div style={{display:'flex',gap:'8px'}}>
        <button onClick={()=>onSave(f)} style={{flex:1,background:PINK,color:'#fff',border:'none',padding:'9px',borderRadius:'8px',cursor:'pointer',fontSize:'12px',fontWeight:600}}>Salvar</button>
        <button onClick={onDelete} style={{background:'transparent',color:'#C0392B',border:'1px solid #C0392B',padding:'9px 14px',borderRadius:'8px',cursor:'pointer',fontSize:'12px'}}>Remover</button>
      </div>
    </div>
  )
}

// ─── Script Card ────────────────────────────────────────────────
function ScriptCard({script,onChange,onDelete}){
  const [copied,setCopied]=useState(false)
  const {t,c,txt}=script
  return (
    <div style={{borderRadius:'8px',overflow:'hidden',border:'0.5px solid var(--border)'}}>
      <div style={{background:c,padding:'10px 14px',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
        <input value={t} onChange={e=>onChange({t:e.target.value})} style={{background:'rgba(255,255,255,.14)',border:'1px solid rgba(255,255,255,.35)',borderRadius:'6px',padding:'5px 8px',fontSize:'12px',fontWeight:600,color:'#fff',width:'min(420px,60%)',outline:'none'}} />
        <div style={{display:'flex',gap:'6px',alignItems:'center'}}>
          <input type="color" value={c} onChange={e=>onChange({c:e.target.value})} title="Cor do script" style={{width:'30px',height:'28px',border:'1px solid rgba(255,255,255,.5)',borderRadius:'6px',background:'transparent',cursor:'pointer',padding:'2px'}} />
          <button onClick={()=>{navigator.clipboard.writeText(txt);setCopied(true);setTimeout(()=>setCopied(false),2000)}} style={{background:'rgba(255,255,255,.2)',border:'none',color:'#fff',padding:'5px 12px',borderRadius:'6px',cursor:'pointer',fontSize:'11px',fontWeight:500}}>{copied?'Copiado!':'Copiar'}</button>
          <button onClick={onDelete} style={{background:'rgba(255,255,255,.12)',border:'1px solid rgba(255,255,255,.35)',color:'#fff',padding:'5px 10px',borderRadius:'6px',cursor:'pointer',fontSize:'11px',fontWeight:500}}>Remover</button>
        </div>
      </div>
      <textarea value={txt} onChange={e=>onChange({txt:e.target.value})} style={{display:'block',width:'100%',minHeight:'150px',background:'var(--surface-1)',border:'none',padding:'12px 14px',fontSize:'12px',lineHeight:1.65,color:'var(--text-secondary)',whiteSpace:'pre-wrap',resize:'vertical',outline:'none',fontFamily:'inherit'}} />
    </div>
  )
}

// ─── Score Badge ────────────────────────────────────────────────
function ScoreBadge({score,small}){
  const c=score>=80?'#1A6B3C':score>=50?'#D4880A':'#C0392B'
  const bg=score>=80?'#e8f5ee':score>=50?'#fdf5e6':'#fdeded'
  const sz=small?22:28
  return <div style={{width:sz,height:sz,borderRadius:'50%',background:bg,display:'flex',alignItems:'center',justifyContent:'center',fontSize:small?9:10,fontWeight:700,color:c}}>{score}</div>
}
