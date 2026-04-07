import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import re
import requests
import time
import unicodedata
import math
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp
import os

# =========================================================
# CONFIGURAÇÃO DA TELA WEB
# =========================================================
st.set_page_config(page_title="Roteirizador Logístico", page_icon="🚛", layout="wide")

# =========================================================
# 1. OS TRADUTORES DE DADOS, HORÁRIOS E NOMES
# =========================================================
def remover_acentos(texto):
    try:
        texto = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
        return texto.lower().strip()
    except:
        return str(texto).lower().strip()

def limpar_nome_cliente(nome):
    nome = remover_acentos(nome)
    termos_remover = [r'\bltda\b', r'\beireli\b', r'\bme\b', r'\bbar\b', r'\brestaurante\b', r'\bcomercio\b', r'\bde\b', r'\be\b', r'\.', r'-']
    for termo in termos_remover:
        nome = re.sub(termo, '', nome)
    return re.sub(r'\s+', ' ', nome).strip()

def limpar_bairro(bairro_cru):
    b = remover_acentos(str(bairro_cru))
    b = b.replace('/ nit', '').replace('/nit', '').replace(' (jacarepagua)', '').strip()
    if b == "freguesia jacarepagua": return "freguesia"
    return b

def limpar_peso(valor_excel):
    try:
        if pd.isna(valor_excel): return 0
        peso_float = float(str(valor_excel).strip().replace(',', '.'))
        if peso_float > 5000: peso_float = peso_float / 1000
        return int(peso_float)
    except: return 0

def parse_hora_str(hora_str):
    if ':' in hora_str:
        h, m = map(int, hora_str.split(':'))
    else:
        h, m = int(hora_str), 0
    if h == 0: h = 24 
    return max(0, ((h * 60) + m) - 480) 

def traduzir_janela(texto_excel):
    limite_maximo = 1440 
    horario_comercial = 600 
    if pd.isna(texto_excel) or str(texto_excel).strip() == '': 
        return [(0, horario_comercial)]
    texto = str(texto_excel).lower()
    janelas = []
    for parte in texto.split('ou'):
        parte = parte.replace(' ', '').replace('h', '')
        tempos = re.findall(r'\d{1,2}:\d{2}|\d{1,2}', parte)
        if not tempos: continue
        if 'partir' in parte or 'apos' in parte or 'após' in parte:
            janelas.append((parse_hora_str(tempos[0]), limite_maximo))
            continue
        if 'até' in parte or 'ate' in parte:
            janelas.append((0, min(limite_maximo, parse_hora_str(tempos[0]))))
            continue
        if len(tempos) >= 2:
            ini, fim = parse_hora_str(tempos[0]), min(limite_maximo, parse_hora_str(tempos[1]))
            if ini > fim: ini, fim = fim, ini
            janelas.append((ini, fim))
        elif len(tempos) == 1:
            janelas.append((0, limite_maximo))
    return janelas if janelas else [(0, horario_comercial)]

def traduzir_hora_inicio(texto_hora):
    try:
        if pd.isna(texto_hora) or str(texto_hora).strip() == '': return 0
        if type(texto_hora).__name__ == 'time':
            return max(0, (texto_hora.hour * 60 + texto_hora.minute) - 480)
        texto = str(texto_hora).lower().replace('h', ':')
        nums = re.findall(r'\d+', texto)
        if len(nums) >= 2: return max(0, (int(nums[0]) * 60 + int(nums[1])) - 480)
        elif len(nums) == 1: return max(0, (int(nums[0]) * 60) - 480)
        return 0
    except: return 0

def traduzir_hora_exata(texto_hora):
    try:
        if pd.isna(texto_hora) or str(texto_hora).strip() == '': return -1
        if type(texto_hora).__name__ == 'time':
            return max(0, (texto_hora.hour * 60 + texto_hora.minute) - 480)
        texto = str(texto_hora).lower().replace('h', ':')
        nums = re.findall(r'\d+', texto)
        if len(nums) >= 2: return max(0, (int(nums[0]) * 60 + int(nums[1])) - 480)
        elif len(nums) == 1: return max(0, (int(nums[0]) * 60) - 480)
        return -1
    except: return -1

# =========================================================
# 2. O MOTOR GEOGRÁFICO & DICIONÁRIOS ATUALIZADOS
# =========================================================
MACRO_ZONAS = {
    'zs': ['botafogo', 'copacabana', 'ipanema', 'leblon', 'flamengo', 'laranjeiras', 'humaita', 'gloria', 'gavea', 'jardim botanico', 'sao conrado', 'lapa', 'vidigal', 'catete', 'urca', 'lagoa', 'leme', 'parque do flamengo'],
    'zn': ['tijuca', 'maracana', 'meier', 'bonsucesso', 'benfica', 'vila isabel', 'olaria', 'maria da graca', 'riachuelo', 'cascadura', 'cachambi', 'andarai', 'engenho de dentro', 'jardim guanabara', 'higienopolis', 'vicente de carvalho', 'todos os santos', 'iraja'],
    'zo': ['barra da tijuca', 'recreio dos bandeirantes', 'recreio', 'jacarepagua', 'freguesia', 'itanhanga', 'joa', 'anil', 'grumari'],
    'centro': ['centro', 'santa teresa', 'saude', 'gamboa'],
    'niteroi': ['niteroi', 'icarai', 'sao francisco', 'piratininga', 'santa rosa', 'varzea das mocas'],
    'baixada': ['sao joao de meriti']
}

COORDENADAS_RJ = {
    'botafogo': (-43.1866, -22.9511), 'copacabana': (-43.1848, -22.9711), 'ipanema': (-43.2056, -22.9844), 'leblon': (-43.2217, -22.9828),
    'centro': (-43.1786, -22.9035), 'laranjeiras': (-43.1853, -22.9333), 'flamengo': (-43.1741, -22.9328), 'tijuca': (-43.2325, -22.9231),
    'barra da tijuca': (-43.3661, -23.0004), 'recreio dos bandeirantes': (-43.4658, -23.0183), 'recreio': (-43.4658, -23.0183), 'jacarepagua': (-43.3503, -22.9686),
    'niteroi': (-43.1167, -22.8833), 'humaita': (-43.1983, -22.9583), 'gloria': (-43.1764, -22.9197), 'gavea': (-43.2289, -22.9792),
    'jardim botanico': (-43.2167, -22.9667), 'maracana': (-43.2281, -22.9133), 'meier': (-43.2800, -22.9011), 'bonsucesso': (-43.2556, -22.8631),
    'benfica': (-43.2428, -22.8894), 'freguesia': (-43.3425, -22.9403), 'itanhanga': (-43.3039, -22.9825), 'lapa': (-43.1811, -22.9144),
    'vila isabel': (-43.2425, -22.9147), 'sao conrado': (-43.2653, -22.9922), 'olaria': (-43.2633, -22.8447), 'icarai': (-43.1075, -22.9056),
    'sao francisco': (-43.0903, -22.9253), 'piratininga': (-43.0647, -22.9492), 'maria da graca': (-43.2611, -22.8858), 'riachuelo': (-43.2569, -22.8994),
    'cascadura': (-43.3236, -22.8864), 'santa teresa': (-43.1872, -22.9272), 'vidigal': (-43.2356, -22.9933), 'santa rosa': (-43.1064, -22.9008),
    'catete': (-43.1764, -22.9275), 'urca': (-43.1667, -22.95), 'lagoa': (-43.2058, -22.9722), 'leme': (-43.1714, -22.9622),
    'cachambi': (-43.2781, -22.8906), 'andarai': (-43.2503, -22.9264), 'engenho de dentro': (-43.2953, -22.8953),
    'jardim guanabara': (-43.2017, -22.8167), 'higienopolis': (-43.2614, -22.8711), 'saude': (-43.1844, -22.8958), 'gamboa': (-43.1933, -22.8972), 
    'joa': (-43.2917, -23.0133), 'sao joao de meriti': (-43.3719, -22.8028),
    'anil': (-43.3364, -22.9567), 'grumari': (-43.5208, -23.0483), 'vicente de carvalho': (-43.3131, -22.8522),
    'todos os santos': (-43.2808, -22.8953), 'iraja': (-43.3267, -22.8278), 'varzea das mocas': (-43.0181, -22.8942), 'parque do flamengo': (-43.1741, -22.9328),
    'barra de guaratiba': (-43.5658, -23.0189),
    'ribeira': (-43.1706, -22.8256),
    'ilha do governador': (-43.2036, -22.8064),
    'bangu': (-43.4644, -22.8756),
    'vila valqueire': (-43.3642, -22.8886),
    'campo grande': (-43.5594, -22.9022),
    'bento ribeiro': (-43.3619, -22.8683),
    'vargem pequena': (-43.4608, -22.9961),
    'sepetiba': (-43.6978, -22.9692),
    'praca da bandeira': (-43.2131, -22.9125)
}

def obter_coordenadas(bairro_exato):
    bairro_limpo = limpar_bairro(bairro_exato.split(',')[0])
    if bairro_limpo in COORDENADAS_RJ: return COORDENADAS_RJ[bairro_limpo][0], COORDENADAS_RJ[bairro_limpo][1]
    try:
        r = requests.get(f"https://nominatim.openstreetmap.org/search?q={bairro_exato}, Rio de Janeiro, Brasil&format=json&limit=1", headers={'User-Agent': 'STPP_Web/1.0'}, timeout=5)
        dados = r.json()
        if dados: 
            time.sleep(1)
            return float(dados[0]['lon']), float(dados[0]['lat'])
    except: pass
    return None, None

def haversine_dist(lon1, lat1, lon2, lat2):
    # Calcula a distância em linha reta na Terra (Fórmula de Haversine)
    R = 6371 # Raio da Terra em km
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = math.sin(dLat/2) * math.sin(dLat/2) + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dLon/2) * math.sin(dLon/2)
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    return R * c

def gerar_matriz_osrm(coordenadas):
    coords_str = ";".join([f"{lon},{lat}" for lon, lat in coordenadas])
    
    # 1. TENTA O SERVIDOR PÚBLICO COM RETENTATIVAS (Timeout de 20s)
    for tentativa in range(2):
        try:
            r = requests.get(f"http://router.project-osrm.org/table/v1/driving/{coords_str}?annotations=duration", timeout=20)
            dados = r.json()
            if dados.get('code') == 'Ok': return [[int((val / 60) * 3) for val in row] for row in dados['durations']]
        except:
            time.sleep(2) # Respira e tenta de novo
            
    # 2. PLANO B (FALLBACK MATEMÁTICO) - Se o OSRM cair, calcula localmente!
    matriz_fallback = []
    for i in range(len(coordenadas)):
        linha = []
        for j in range(len(coordenadas)):
            if i == j:
                linha.append(0)
            else:
                dist_km = haversine_dist(coordenadas[i][0], coordenadas[i][1], coordenadas[j][0], coordenadas[j][1])
                # Multiplica por fator 1.5 (curvas de ruas) e considera vel média de 25km/h no RJ (2.4 min/km)
                tempo_estimado = int((dist_km * 1.5) * 2.4)
                linha.append(tempo_estimado)
        matriz_fallback.append(linha)
    
    # Adicionamos um aviso de que o plano B foi ativado
    st.session_state.aviso_fallback = True
    return matriz_fallback

# =========================================================
# 3. CONSTRUÇÃO DO MODELO & PROCESSAMENTO
# =========================================================
def processar_rotas(arquivo_excel):
    try:
        df_pedidos = pd.read_excel(arquivo_excel, sheet_name='Pedidos')
        df_pedidos.columns = df_pedidos.columns.str.strip().str.lower()
        df_frota = pd.read_excel(arquivo_excel, sheet_name='Frota')
        df_frota.columns = df_frota.columns.str.strip().str.lower()
    except Exception as e: return None, f"Erro ao ler abas: {str(e)}", []
        
    df_frota['trabalha hoje'] = df_frota['trabalha hoje'].astype(str).str.strip().str.upper()
    frota_ativa = df_frota[df_frota['trabalha hoje'] == 'SIM'].copy()
    if len(frota_ativa) == 0: return None, "Nenhum motorista ativo.", []

    lon_base, lat_base = -43.2163, -22.9123
    puxadas = []
    try:
        df_puxadas = pd.read_excel(arquivo_excel, sheet_name='Puxadas')
        df_puxadas.columns = df_puxadas.columns.str.strip().str.lower()
        col_nome = 'horário' if 'horário' in df_puxadas.columns else ('horario' if 'horario' in df_puxadas.columns else None)
        if col_nome:
            for _, row in df_puxadas.iterrows():
                minutos = traduzir_hora_exata(row[col_nome])
                if minutos != -1: puxadas.append((minutos, minutos + 60)) 
    except: pass 

    bairros_unicos = df_pedidos['bairro'].astype(str).str.strip().str.title().unique().tolist()
    mapa_indices = {'BASE': 0}
    coordenadas = [(lon_base, lat_base)]
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, bairro in enumerate(bairros_unicos):
        status_text.text(f"📍 Mapeando bairro via GPS: {bairro}")
        blon, blat = obter_coordenadas(bairro)
        
        # Correção ativada: Afasta coordenadas matematicamente se não achar o bairro
        if not blon: 
            blon = lon_base + (0.01 * (i + 1))
            blat = lat_base + (0.01 * (i + 1))
            
        coordenadas.append((blon, blat))
        mapa_indices[limpar_bairro(bairro)] = i + 1
        progress_bar.progress((i + 1) / len(bairros_unicos) * 0.4)

    status_text.text("🚦 Calculando rotas de trânsito...")
    st.session_state.aviso_fallback = False
    matriz_tempos_real = gerar_matriz_osrm(coordenadas)
    
    if not matriz_tempos_real: 
        return None, "Erro crítico ao tentar calcular as distâncias matemáticas.", []
    
    progress_bar.progress(0.6)

    data = {'motoristas': [], 'veiculos': [], 'vehicle_capacities': [], 'vehicle_costs': [], 'vehicle_start_times': [], 'vehicle_preferences': []}
    
    for _, row in frota_ativa.iterrows():
        mot = str(row['motorista']).title()
        veic = str(row['veiculo']).title()
        cap = int(pd.to_numeric(row['capacidade'], errors='coerce'))
        start_t = traduzir_hora_inicio(row['inicio'])
        
        is_prioridade = int(pd.to_numeric(row.get('prioridade', 1), errors='coerce')) == 1
        custo_v1 = 10000 if is_prioridade else 50000
        
        prefs = set()
        col_pref = 'preferência' if 'preferência' in frota_ativa.columns else 'preferencia' if 'preferencia' in frota_ativa.columns else None
        if col_pref and str(row[col_pref]).strip() != 'nan':
            for z in remover_acentos(str(row[col_pref])).split(','):
                z = ''.join([i for i in z if not i.isdigit()]).strip()
                if z in MACRO_ZONAS:
                    for bm in MACRO_ZONAS[z]:
                        if bm in mapa_indices: prefs.add(mapa_indices[bm])

        data['motoristas'].append(f"{mot} [Viagem 1]")
        data['veiculos'].append(veic)
        data['vehicle_capacities'].append(cap)
        data['vehicle_start_times'].append(start_t)
        data['vehicle_costs'].append(custo_v1)
        data['vehicle_preferences'].append(prefs)

        data['motoristas'].append(f"{mot} [Viagem 2]")
        data['veiculos'].append(veic)
        data['vehicle_capacities'].append(cap)
        data['vehicle_start_times'].append(start_t)
        data['vehicle_costs'].append(custo_v1 + 500000) 
        data['vehicle_preferences'].append(prefs)

    data['num_vehicles'] = len(data['motoristas'])
    data['puxadas'] = puxadas
    data['time_matrix'] = matriz_tempos_real
    data['locations'] = [0]
    data['bairros_exatos'] = ['BASE'] 
    data['zonas_exatas'] = ['BASE']
    data['demands'] = [0]
    data['pesos_reais'] = [0]
    data['time_windows'] = [[(0, 1440)]]
    data['service_time'] = [45] 
    data['nomes_clientes'] = ['BASE']
    data['nomes_clientes_limpos'] = ['BASE'] 
    data['nfs'] = ['-']
    data['cervejarias'] = ['-'] 
    data['tipos'] = ['BASE']
    data['coordenadas'] = coordenadas 
    
    clientes_vistos = [] 
    for _, row in df_pedidos.iterrows():
        bairro_cru = str(row['bairro'])
        bairro_limpo = limpar_bairro(bairro_cru)
        peso = limpar_peso(row['peso'])
        tipo = str(row['tipo']).strip().upper() if 'tipo' in df_pedidos.columns and not pd.isna(row['tipo']) else 'ENTREGA'
        cliente_cru = str(row['cliente']).strip()
        cliente_limpo = limpar_nome_cliente(cliente_cru)
        
        ja_visto = False
        for c_visto, b_visto in clientes_vistos:
            if bairro_limpo == b_visto and (cliente_limpo == c_visto or cliente_limpo in c_visto or c_visto in cliente_limpo):
                ja_visto = True; break
        
        tempo_servico = 2 if ja_visto else 15
        if not ja_visto: clientes_vistos.append((cliente_limpo, bairro_limpo))
            
        macro_do_pedido = 'desconhecida'
        for mz, lista_bairros in MACRO_ZONAS.items():
            if bairro_limpo in lista_bairros: macro_do_pedido = mz; break
        
        data['demands'].append(peso if tipo == 'COLETA' else -peso)
        data['locations'].append(mapa_indices.get(bairro_limpo, 0)) 
        data['bairros_exatos'].append(bairro_cru)
        data['zonas_exatas'].append(macro_do_pedido)
        data['pesos_reais'].append(peso)
        data['tipos'].append(tipo)
        data['time_windows'].append(traduzir_janela(row['janela']))
        data['service_time'].append(tempo_servico) 
        data['nomes_clientes'].append(cliente_cru)
        data['nomes_clientes_limpos'].append(cliente_limpo)
        data['nfs'].append(str(row['nf']).strip() if 'nf' in df_pedidos.columns and not pd.isna(row['nf']) else 'S/N')
        data['cervejarias'].append(str(row['cervejaria']).strip() if 'cervejaria' in df_pedidos.columns and not pd.isna(row['cervejaria']) else '-')
    
    data['depot'] = 0

    status_text.text("🧠 Otimizando a matemática de rotas...")
    manager = pywrapcp.RoutingIndexManager(len(data['time_windows']), data['num_vehicles'], data['depot'])
    routing = pywrapcp.RoutingModel(manager)

    def calc_tempo_real(from_n, to_n):
        if from_n == data['depot'] or to_n == data['depot']: return data['time_matrix'][data['locations'][from_n]][data['locations'][to_n]]
        c1, c2 = data['nomes_clientes_limpos'][from_n], data['nomes_clientes_limpos'][to_n]
        b1, b2 = data['bairros_exatos'][from_n], data['bairros_exatos'][to_n]
        if (c1 == c2 or c1 in c2 or c2 in c1) and b1 == b2: return 0
        if b1 == b2: return 5
        return data['time_matrix'][data['locations'][from_n]][data['locations'][to_n]]

    def calc_tempo_horario(from_n, to_n):
        if to_n == data['depot']: return 0
        return calc_tempo_real(from_n, to_n)

    def time_cb(from_idx, to_idx): return calc_tempo_horario(manager.IndexToNode(from_idx), manager.IndexToNode(to_idx)) + data['service_time'][manager.IndexToNode(from_idx)]
    time_cb_idx = routing.RegisterTransitCallback(time_cb)
    
    routing.AddDimension(time_cb_idx, 2880, 2880, False, 'Time')
    time_dim = routing.GetDimensionOrDie('Time')

    def create_cost_cb(v_id):
        def cost_cb(f_idx, t_idx):
            f_n, t_n = manager.IndexToNode(f_idx), manager.IndexToNode(t_idx)
            c = calc_tempo_real(f_n, t_n) + data['service_time'][f_n]
            
            if t_n != data['depot'] and data['vehicle_preferences'][v_id] and data['locations'][t_n] not in data['vehicle_preferences'][v_id]: 
                c += 180  
                
            if f_n != data['depot'] and t_n != data['depot']:
                zona_origem = data['zonas_exatas'][f_n]
                zona_destino = data['zonas_exatas'][t_n]
                if zona_origem != 'desconhecida' and zona_destino != 'desconhecida' and zona_origem != zona_destino: 
                    c += 120 
            return c
        return cost_cb

    for v in range(data['num_vehicles']):
        cb_idx = routing.RegisterTransitCallback(create_cost_cb(v))
        routing.SetArcCostEvaluatorOfVehicle(cb_idx, v)
        routing.SetFixedCostOfVehicle(data['vehicle_costs'][v], v)

    def demand_cb(f_idx): return data['demands'][manager.IndexToNode(f_idx)]
    demand_cb_idx = routing.RegisterUnaryTransitCallback(demand_cb)
    routing.AddDimensionWithVehicleCapacity(demand_cb_idx, 0, data['vehicle_capacities'], False, 'Capacity')

    penalty = 10000000 
    for node in range(1, len(data['time_windows'])): routing.AddDisjunction([manager.NodeToIndex(node)], penalty)

    for loc_idx, janelas in enumerate(data['time_windows']):
        if loc_idx == data['depot']: continue
        idx = manager.NodeToIndex(loc_idx)
        var = time_dim.CumulVar(idx)
        if len(janelas) == 1: var.SetRange(janelas[0][0], janelas[0][1])
        else:
            j_sort = sorted(janelas, key=lambda x: x[0])
            var.SetRange(j_sort[0][0], j_sort[-1][1]) 
            for i in range(len(j_sort) - 1):
                if j_sort[i+1][0] > j_sort[i][1]: var.RemoveInterval(j_sort[i][1] + 1, j_sort[i+1][0] - 1)
        
    for v in range(data['num_vehicles']):
        idx_saida = routing.Start(v)
        time_dim.CumulVar(idx_saida).SetRange(int(data['vehicle_start_times'][v]), 2880)
        routing.AddVariableMinimizedByFinalizer(time_dim.CumulVar(idx_saida))
        routing.AddVariableMinimizedByFinalizer(time_dim.CumulVar(routing.End(v)))

    solver = routing.solver()

    num_carros_reais = len(frota_ativa)
    for i in range(num_carros_reais):
        v1_idx = i * 2       
        v2_idx = i * 2 + 1   
        
        end_v1 = time_dim.CumulVar(routing.End(v1_idx))
        start_v2 = time_dim.CumulVar(routing.Start(v2_idx))
        
        solver.Add(start_v2 >= end_v1 + 45)

    intervalos_doca = []
    run_id = str(int(time.time() * 1000)) 
    
    for v in range(data['num_vehicles']):
        start_v = time_dim.CumulVar(routing.Start(v))
        is_active_v = routing.NextVar(routing.Start(v)) != routing.End(v)
        
        inicio_carregamento = solver.Sum([start_v, -45 * is_active_v])
        intervalo_opcional = solver.FixedDurationIntervalVar(0, 2880, 45, True, f"carregamento_opt_{v}_{run_id}")
        
        solver.Add(intervalo_opcional.PerformedExpr() == is_active_v)
        solver.Add(intervalo_opcional.StartExpr() == inicio_carregamento.Var())

        intervalos_doca.append(intervalo_opcional)

        for p_start, p_end in data['puxadas']:
            solver.Add(start_start <= (p_start - 45) + start_v >= p_end + (1 - is_active_v) >= 1)

    solver.Add(solver.DisjunctiveConstraint(intervalos_doca, f"Fila_da_Doca_{run_id}"))

    params = pywrapcp.DefaultRoutingSearchParameters()
    params.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PARALLEL_CHEAPEST_INSERTION
    params.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    params.time_limit.seconds = 30 
    
    progress_bar.progress(0.8)
    sol = routing.SolveWithParameters(params)

    if sol:
        status_text.text("✅ Rotas finalizadas! Gerando relatórios e mapa...")
        dados_excel = []
        dropped_nodes_info = []
        mapa_linhas = []

        for node in range(routing.Size()):
            if routing.IsStart(node) or routing.IsEnd(node): continue
            if sol.Value(routing.NextVar(node)) == node:
                idx = manager.IndexToNode(node)
                dropped_nodes_info.append({
                    'NF': str(data['nfs'][idx]).split('.')[0],
                    'Cliente': data['nomes_clientes'][idx].title(),
                    'Bairro': data['bairros_exatos'][idx].title(),
                    'Peso (kg)': data['pesos_reais'][idx]
                })

        cores = ['blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue', 'darkgreen', 'black']
        
        for vehicle_id in range(data['num_vehicles']):
            index = routing.Start(vehicle_id)
            if routing.IsEnd(sol.Value(routing.NextVar(index))): continue 
            
            nome_motorista = data['motoristas'][vehicle_id]
            carro = data['veiculos'][vehicle_id]
            carga_atual = 0
            temp_idx = index
            while not routing.IsEnd(temp_idx):
                n_idx = manager.IndexToNode(temp_idx)
                if data['tipos'][n_idx] == 'ENTREGA': carga_atual += data['pesos_reais'][n_idx]
                temp_idx = sol.Value(routing.NextVar(temp_idx))
                
            rota_coords = []
            
            while not routing.IsEnd(index):
                time_var = time_dim.CumulVar(index)
                n_idx = manager.IndexToNode(index)
                min_totais = int((8 * 60) + sol.Min(time_var))
                hora = int((min_totais // 60) % 24)
                minuto = int(min_totais % 60)
                
                lon_lat = data['coordenadas'][data['locations'][n_idx]]
                rota_coords.append((lon_lat[1], lon_lat[0]))

                if n_idx == 0: 
                    hora_partida = hora
                    min_partida = minuto
                    dados_excel.append({'Motorista / Veículo': f"{nome_motorista} ({carro})", 'Horário': f"{hora_partida:02d}:{min_partida:02d}", 'Ação': 'SAÍDA DA BASE', 'NF': '-', 'Cervejaria': '-', 'Cliente': 'BASE DA EMPRESA', 'Bairro': 'BASE', 'Peso (kg)': f"{carga_atual} (Total Carregado)"})
                else:
                    peso = data['pesos_reais'][n_idx]
                    tipo = data['tipos'][n_idx]
                    if tipo == 'ENTREGA': carga_atual -= peso
                    else: carga_atual += peso
                    dados_excel.append({'Motorista / Veículo': f"{nome_motorista} ({carro})", 'Horário': f"{hora:02d}:{minuto:02d}", 'Ação': tipo, 'NF': str(data['nfs'][n_idx]).split('.')[0], 'Cervejaria': data['cervejarias'][n_idx], 'Cliente': data['nomes_clientes'][n_idx].title(), 'Bairro': data['bairros_exatos'][n_idx].title(), 'Peso (kg)': peso})
                index = sol.Value(routing.NextVar(index))
                
            time_var = time_dim.CumulVar(index)
            min_totais = int((8 * 60) + sol.Min(time_var))
            hora_fim = int((min_totais // 60) % 24)
            min_fim = int(min_totais % 60)
            dados_excel.append({'Motorista / Veículo': f"{nome_motorista} ({carro})", 'Horário': f"{hora_fim:02d}:{min_fim:02d}", 'Ação': 'FIM DA VIAGEM', 'NF': '-', 'Cervejaria': '-', 'Cliente': 'RETORNO À BASE', 'Bairro': '-', 'Peso (kg)': f"{carga_atual} (Vazios)"})
            dados_excel.append({k: "" for k in dados_excel[0].keys()})
            
            if rota_coords:
                mapa_linhas.append({"motorista": nome_motorista, "coords": rota_coords, "cor": cores[vehicle_id % len(cores)]})

        if dropped_nodes_info:
            if dados_excel: dados_excel.append({k: "" for k in dados_excel[0].keys()}) 
            dados_excel.append({'Motorista / Veículo': '--- NOTAS CORTADAS ---', 'Horário': '---', 'Ação': '---', 'NF': '---', 'Cervejaria': '---', 'Cliente': '---', 'Bairro': '---', 'Peso (kg)': '---'})
            for drop in dropped_nodes_info:
                dados_excel.append({
                    'Motorista / Veículo': '⚠️ NÃO ALOCADO', 'Horário': '-', 'Ação': 'CORTADA / NA BASE',
                    'NF': drop['NF'], 'Cervejaria': '-', 'Cliente': drop['Cliente'], 'Bairro': drop['Bairro'], 'Peso (kg)': drop['Peso (kg)']
                })
        
        progress_bar.progress(1.0)
        status_text.empty()
        progress_bar.empty()
        return pd.DataFrame(dados_excel), mapa_linhas, dropped_nodes_info
    else: 
        return None, "Impossível achar uma rota com as regras atuais.", []

# =========================================================
# FRONTEND / INTERFACE DE USUÁRIO (UI)
# =========================================================

col1, col2, col3 = st.columns([3, 1, 3])
with col2:
    try:
        from PIL import Image
        img = Image.open("logo.png")
        st.image(img, use_container_width=True)
    except Exception as e:
        pass

st.markdown("<h1 style='text-align: center;'>Roteirizador Web</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Carregue a planilha de <b>pedidos e frota</b> para desenhar as rotas automaticamente.</p>", unsafe_allow_html=True)

arquivo_upload = st.file_uploader("Arraste a sua planilha Excel (romaneio_teste.xlsx) aqui", type=["xlsx"])

if 'rotas_geradas' not in st.session_state:
    st.session_state.rotas_geradas = False

if arquivo_upload is not None:
    if st.button("🚀 Gerar Rotas Otimizadas"):
        with st.spinner("Iniciando o Cérebro Logístico..."):
            df_resultado, linhas_mapa, notas_cortadas = processar_rotas(arquivo_upload)
            
            if df_resultado is not None and not isinstance(df_resultado, str):
                st.session_state.rotas_geradas = True
                st.session_state.df_resultado = df_resultado
                st.session_state.linhas_mapa = linhas_mapa
                st.session_state.notas_cortadas = notas_cortadas
            else:
                st.error(f"Ocorreu um erro no cálculo: {linhas_mapa}")

    if st.session_state.rotas_geradas:
        st.success("Roteamento concluído com sucesso!")
        
        if getattr(st.session_state, 'aviso_fallback', False):
            st.warning("⚠️ **ATENÇÃO:** O servidor de trânsito público (OSRM) não respondeu a tempo. Para evitar a paralisação da operação, as distâncias foram estimadas via fórmula matemática (Haversine). Os tempos podem apresentar pequenas variações da realidade.")

        if st.session_state.notas_cortadas:
            st.error(f"🚨 ATENÇÃO: {len(st.session_state.notas_cortadas)} notas foram deixadas na base devido ao limite de tempo ou capacidade!")
            df_cortadas = pd.DataFrame(st.session_state.notas_cortadas)
            st.table(df_cortadas) 
        else:
            st.success("🎉 Excelente notícia: 100% das notas foram roteirizadas! Nenhuma carga ficou para trás na Doca.")
            
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📊 Tabela de Viagem")
            st.dataframe(st.session_state.df_resultado, use_container_width=True, height=500)
            
            csv = st.session_state.df_resultado.to_csv(index=False).encode('utf-8')
            st.download_button(label="📥 Baixar Planilha de Roteamento (CSV)", data=csv, file_name='Rotas_Geradas.csv', mime='text/csv')
        
        with col2:
            st.subheader("🗺️ Mapa Interativo de Rotas")
            m = folium.Map(location=[-22.9123, -43.2163], zoom_start=11)
            
            folium.Marker([-22.9123, -43.2163], popup="🏠 BASE STPP", icon=folium.Icon(color="black", icon="home")).add_to(m)
            
            for linha in st.session_state.linhas_mapa:
                folium.PolyLine(
                    locations=linha["coords"],
                    color=linha["cor"],
                    weight=5,
                    opacity=0.8,
                    tooltip=f"Rota: {linha['motorista']}"
                ).add_to(m)
            
            st_folium(m, width=700, height=500)
else:
    st.session_state.rotas_geradas = False
