import streamlit as st
import psycopg2
import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import os
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA STREAMLIT ---
st.set_page_config(page_title="Prode Mundial 2026", page_icon="⚽", layout="wide")

# --- CONEXIÓN DE ARQUITECTURA PROFESIONAL (POSTGRESQL CLOUD) ---
class DatabaseManager:
    @staticmethod
    def get_connection():
        return psycopg2.connect(st.secrets["postgres"]["url"])

    @staticmethod
    def init_db():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        
        comandos_tablas = [
            '''CREATE TABLE IF NOT EXISTS equipos (id_equipo SERIAL PRIMARY KEY, nombre TEXT UNIQUE NOT NULL, zona TEXT NOT NULL, archivo_bandera TEXT DEFAULT 'default.png')''',
            '''CREATE TABLE IF NOT EXISTS partidos (id_partido SERIAL PRIMARY KEY, fase TEXT NOT NULL, id_equipo_local INTEGER, id_equipo_visitante INTEGER, goles_local INTEGER DEFAULT NULL, goles_visitante INTEGER DEFAULT NULL, FOREIGN KEY(id_equipo_local) REFERENCES equipos(id_equipo), FOREIGN KEY(id_equipo_visitante) REFERENCES equipos(id_equipo), UNIQUE(fase, id_equipo_local, id_equipo_visitante))''',
            '''CREATE TABLE IF NOT EXISTS usuarios (id_usuario SERIAL PRIMARY KEY, nombre TEXT UNIQUE NOT NULL)''',
            '''CREATE TABLE IF NOT EXISTS apuestas (id_apuesta SERIAL PRIMARY KEY, id_usuario INTEGER, id_partido INTEGER, apuesta_goles_local INTEGER DEFAULT NULL, apuesta_goles_visitante INTEGER DEFAULT NULL, equipo_l_predicho TEXT, equipo_v_predicho TEXT, puntos_obtenidos REAL DEFAULT 0.0, FOREIGN KEY(id_usuario) REFERENCES usuarios(id_usuario), FOREIGN KEY(id_partido) REFERENCES partidos(id_partido), UNIQUE(id_usuario, id_partido))''',
            '''CREATE TABLE IF NOT EXISTS configuracion (id INTEGER PRIMARY KEY, pts_prode INTEGER, pts_exacto INTEGER, pts_parcial INTEGER, pts_dif INTEGER, api_key TEXT, id_liga TEXT, admin_pass TEXT, pts_prode_ko INTEGER DEFAULT 8, pts_exacto_ko INTEGER DEFAULT 6, pts_parcial_ko INTEGER DEFAULT 2, pts_dif_ko INTEGER DEFAULT 2, fecha_limite TEXT DEFAULT '2026-06-11 16:00:00')'''
        ]
        
        for cmd in comandos_tablas:
            try:
                cursor = conn.cursor()
                cursor.execute(cmd)
                conn.commit()
                cursor.close()
            except Exception:
                conn.rollback()
        
        # --- MIGRADOR AUTOMÁTICO DE FIXTURE ---
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM equipos WHERE nombre = 'Costa Rica'")
            if cursor.fetchone()[0] > 0:
                cursor.execute("TRUNCATE apuestas, partidos, equipos, usuarios RESTART IDENTITY CASCADE;")
                conn.commit()
            cursor.close()
        except Exception:
            conn.rollback()

        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO configuracion (id, pts_prode, pts_exacto, pts_parcial, pts_dif, api_key, id_liga, admin_pass, pts_prode_ko, pts_exacto_ko, pts_parcial_ko, pts_dif_ko, fecha_limite) 
                VALUES (1, 4, 3, 1, 1, '', '1', 'admin123', 8, 6, 2, 2, '2026-06-11 16:00:00')
                ON CONFLICT (id) DO NOTHING
            """)
            conn.commit()
            cursor.close()
        except Exception:
            conn.rollback()
        
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM equipos")
            if cursor.fetchone()[0] == 0:
                # MAPA OFICIAL DE TUS 12 GRUPOS
                mapa_grupos = {
                    "A": ["México", "Sudáfrica", "Corea del Sur", "Rep. Checa"],
                    "B": ["Canadá", "Bosnia", "Qatar", "Suiza"],
                    "C": ["Brasil", "Marruecos", "Haití", "Escocia"],
                    "D": ["EE.UU.", "Paraguay", "Australia", "Turquía"],
                    "E": ["Alemania", "Curazao", "C. Marfil", "Ecuador"],
                    "F": ["Países Bajos", "Japón", "Suecia", "Túnez"],
                    "G": ["Bélgica", "Egipto", "Irán", "N. Zelanda"],
                    "H": ["España", "C. Verde", "A. Saudita", "Uruguay"],
                    "I": ["Francia", "Senegal", "Irak", "Noruega"],
                    "J": ["Argentina", "Argelia", "Austria", "Jordania"],
                    "K": ["Portugal", "RD Congo", "Uzbekistán", "Colombia"],
                    "L": ["Inglaterra", "Croacia", "Ghana", "Panamá"]
                }
                iso_mapping = {
                    "México": "mx", "Sudáfrica": "za", "Corea del Sur": "kr", "Rep. Checa": "cz",
                    "Canadá": "ca", "Bosnia": "ba", "Qatar": "qa", "Suiza": "ch",
                    "Brasil": "br", "Marruecos": "ma", "Haití": "ht", "Escocia": "gb-sct",
                    "EE.UU.": "us", "Paraguay": "py", "Australia": "au", "Turquía": "tr",
                    "Alemania": "de", "Curazao": "cw", "C. Marfil": "ci", "Ecuador": "ec",
                    "Países Bajos": "nl", "Japón": "jp", "Suecia": "se", "Túnez": "tn",
                    "Bélgica": "be", "Egipto": "eg", "Irán": "ir", "N. Zelanda": "nz",
                    "España": "es", "C. Verde": "cv", "A. Saudita": "sa", "Uruguay": "uy",
                    "Francia": "fr", "Senegal": "sn", "Irak": "iq", "Noruega": "no",
                    "Argentina": "ar", "Argelia": "dz", "Austria": "at", "Jordania": "jo",
                    "Portugal": "pt", "RD Congo": "cd", "Uzbekistán": "uz", "Colombia": "co",
                    "Inglaterra": "gb-eng", "Croacia": "hr", "Ghana": "gh", "Panamá": "pa"
                }
                for grupo, selecciones in mapa_grupos.items():
                    for sel in selecciones:
                        iso = iso_mapping.get(sel, "un")
                        cursor.execute('INSERT INTO equipos (nombre, zona, archivo_bandera) VALUES (%s, %s, %s) ON CONFLICT (nombre) DO NOTHING', (sel, grupo, f"https://flagcdn.com/w40/{iso}.png"))
                
                cursor.execute("SELECT id_equipo, zona FROM equipos ORDER BY zona, id_equipo")
                eqs = cursor.fetchall()
                zonas = {}
                for eq in eqs: zonas.setdefault(eq[1], []).append(eq[0])
                partidos = []
                for zona, ids in zonas.items():
                    if len(ids) == 4:
                        cruces = [(ids[0], ids[1]), (ids[2], ids[3]), (ids[0], ids[2]), (ids[1], ids[3]), (ids[0], ids[3]), (ids[1], ids[2])]
                        for l, v in cruces: partidos.append((f"Grupo {zona}", l, v))
                
                query_insert_partidos = '''
                    INSERT INTO partidos (fase, id_equipo_local, id_equipo_visitante) 
                    VALUES (%s, %s, %s) 
                    ON CONFLICT (fase, id_equipo_local, id_equipo_visitante) DO NOTHING
                '''
                cursor.executemany(query_insert_partidos, partidos)
                
            conn.commit()
            cursor.close()
        except Exception:
            conn.rollback()
            
        conn.close()

    @staticmethod
    def get_config():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT pts_prode, pts_exacto, pts_parcial, pts_dif, api_key, id_liga, admin_pass, pts_prode_ko, pts_exacto_ko, pts_parcial_ko, pts_dif_ko, fecha_limite FROM configuracion WHERE id = 1")
        datos = cursor.fetchone()
        conn.close()
        return datos

    @staticmethod
    def set_config(p_prode, p_exact, p_parcial, p_dif, api_key, id_liga, admin_pass, p_prode_ko, p_exact_ko, p_parcial_ko, p_dif_ko, fecha_limite):
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE configuracion SET pts_prode=%s, pts_exacto=%s, pts_parcial=%s, pts_dif=%s, api_key=%s, id_liga=%s, admin_pass=%s, pts_prode_ko=%s, pts_exacto_ko=%s, pts_parcial_ko=%s, pts_dif_ko=%s, fecha_limite=%s WHERE id=1", (p_prode, p_exact, p_parcial, p_dif, api_key, id_liga, admin_pass, p_prode_ko, p_exact_ko, p_parcial_ko, p_dif_ko, fecha_limite))
        conn.commit()
        conn.close()

    @staticmethod
    def get_equipos_lista():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id_equipo, nombre FROM equipos ORDER BY nombre")
        datos = cursor.fetchall()
        conn.close()
        return datos

    @staticmethod
    def get_partidos_con_nombres():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT 
                p.id_partido as "id_partido", p.fase as "Fase", el.archivo_bandera as "bandera_l", el.nombre as "Local", 
                p.goles_local as "GL Real", p.goles_visitante as "GV Real", ev.nombre as "Visitante", ev.archivo_bandera as "bandera_v"
            FROM partidos p 
            JOIN equipos el ON p.id_equipo_local = el.id_equipo 
            JOIN equipos ev ON p.id_equipo_visitante = ev.id_equipo 
            ORDER BY 
                CASE 
                    WHEN p.fase LIKE 'Grupo%%' THEN 1 WHEN p.fase = '16avos' THEN 2 WHEN p.fase = '8vos' THEN 3
                    WHEN p.fase = '4tos' THEN 4 WHEN p.fase = 'Semi' THEN 5 WHEN p.fase = 'Final' THEN 6 ELSE 7 
                END, p.id_partido
        '''
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(cursor.fetchall(), columns=columns)
        conn.close()
        return df

    @staticmethod
    def guardar_resultado(id_partido, goles_l, goles_v):
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE partidos SET goles_local = %s, goles_visitante = %s WHERE id_partido = %s", (goles_l, goles_v, id_partido))
        conn.commit()
        conn.close()

    @staticmethod
    def importar_resultados_excel_admin(df):
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            try:
                id_partido = int(row['ID_Partido'])
                if pd.isna(row['Goles_L']) or pd.isna(row['Goles_V']): continue
                cursor.execute("UPDATE partidos SET goles_local = %s, goles_visitante = %s WHERE id_partido = %s", (int(row['Goles_L']), int(row['Goles_V']), id_partido))
            except: continue
        conn.commit()
        conn.close()

    @staticmethod
    def get_usuarios():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT nombre FROM usuarios ORDER BY nombre")
        datos = [r[0] for r in cursor.fetchall()]
        conn.close()
        return datos

    @staticmethod
    def importar_apuestas_excel(df):
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            try: id_partido = int(row['ID_Partido'])
            except: continue
            competidor = str(row['Competidor']).strip()
            cursor.execute("INSERT INTO usuarios (nombre) VALUES (%s) ON CONFLICT (nombre) DO NOTHING", (competidor,))
            cursor.execute("SELECT id_usuario FROM usuarios WHERE nombre = %s", (competidor,))
            id_usr = cursor.fetchone()[0]
            eq_l, eq_v = str(row['Local']).strip(), str(row['Visitante']).strip()
            try: goles_l, goles_v = int(row['Goles_L']), int(row['Goles_V'])
            except: continue
            cursor.execute('''INSERT INTO apuestas (id_usuario, id_partido, apuesta_goles_local, apuesta_goles_visitante, equipo_l_predicho, equipo_v_predicho) VALUES (%s, %s, %s, %s, %s, %s) ON CONFLICT(id_usuario, id_partido) DO UPDATE SET apuesta_goles_local=EXCLUDED.apuesta_goles_local, apuesta_goles_visitante=EXCLUDED.apuesta_goles_visitante, equipo_l_predicho=EXCLUDED.equipo_l_predicho, equipo_v_predicho=EXCLUDED.equipo_v_predicho''', (id_usr, id_partido, goles_l, goles_v, eq_l, eq_v))
        conn.commit()
        conn.close()

    @staticmethod
    def calcular_ranking_avanzado():
        conn = DatabaseManager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT pts_prode, pts_exacto, pts_parcial, pts_dif, pts_prode_ko, pts_exacto_ko, pts_parcial_ko, pts_dif_ko FROM configuracion WHERE id = 1")
        cfg_pts = cursor.fetchone()
        
        cursor.execute('''SELECT p.id_partido, p.fase, el.nombre, ev.nombre, p.goles_local, p.goles_visitante FROM partidos p JOIN equipos el ON p.id_equipo_local = el.id_equipo JOIN equipos ev ON p.id_equipo_visitante = ev.id_equipo WHERE p.goles_local IS NOT NULL''')
        partidos_jugados = cursor.fetchall()
        
        for id_partido, fase, real_l_nombre, real_
