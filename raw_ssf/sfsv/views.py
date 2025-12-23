
from django.db import connection
from django.shortcuts import render

def homepage(request):

	tipo = request.GET.get('tipo', 'todos')

	if tipo == 'todos':
		sql = """
			SELECT
				Entidad,
				Tipo,
				round(sum(Saldo)/1000,1) as Saldo
			FROM
				carteras_clean
			WHERE
				mes = (select max(mes) from carteras_clean)
			GROUP BY
				Entidad,
				Tipo
			ORDER BY
				round(sum(Saldo)/1000,1) desc
		"""
	else:
		sql = f"""
			SELECT
				Entidad,
				Tipo,
				round(sum(Saldo)/1000,1) as Saldo
			FROM
				carteras_clean
			WHERE
				mes = (select max(mes) from carteras_clean)
				and Tipo = '{tipo}'
			GROUP BY
				Entidad,
				Tipo
			ORDER BY
				round(sum(Saldo)/1000,1) desc
		"""


	with connection.cursor() as cursor:
		cursor.execute(sql)
		rows = cursor.fetchall()
	
	labels = [row[0] for row in rows]
	values = [row[2] for row in rows]

	return render(request,'inicio.html',{'labels':labels,'values':values})


def crecimiento(request):

	institucion = request.GET.get('institucion', 'todos')

	str_q = "WHERE 1 = 1" if institucion == 'todos' else f"WHERE Entidad = '{institucion}'"


	sql = f"""
		WITH base AS (
		    SELECT
		        mes,
		        SUM(saldo) AS saldo
		    FROM carteras_clean
		    {str_q}
		    GROUP BY mes
		),
		calc AS (
		    SELECT
		        mes,
		        saldo,
		        LAG(saldo) OVER (ORDER BY mes) AS saldo_anterior,
		        FIRST_VALUE(saldo) OVER (
		            ORDER BY mes
		            ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
		        ) AS saldo_inicial
		    FROM base
		)
		SELECT
		    mes,
		    saldo,
		    saldo_anterior,
		    ROUND(
		        (saldo - saldo_anterior) * 100.0 / saldo_anterior,
		        2
		    ) AS variacion_pct_mes,
		    ROUND(
		        (saldo - saldo_inicial) * 100.0 / saldo_inicial,
		        2
		    ) AS crecimiento_pct_acumulado
		FROM calc
		WHERE mes > '2025-01-31'
		ORDER BY mes;
	"""



	with connection.cursor() as cursor:
		cursor.execute(sql)
		rows = cursor.fetchall()
	
	labels = [row[0] for row in rows]
	values = [row[3] for row in rows]
	values_acc = [row[4] for row in rows]

	return render(request,'crecimiento.html',{'labels':labels,'values':values,'values_acc':values_acc})
