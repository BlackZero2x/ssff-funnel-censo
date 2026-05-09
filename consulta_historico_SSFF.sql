SELECT 
     a.ccod_cli
	,a.[mes]
    ,a.[ccod_ruta]
	,b.[categoria]
	,b.linea
	,b.sublinea
	,b.producto
	,a.[fecha]
	,SUM(a.[cantidad]) AS [pedidos]
    ,SUM(a.[volumen]) AS [total_volumen]
    ,SUM(a.[monto]) AS [total_monto]
FROM [eAuren].[dbo].[base_com] a
LEFT JOIN [comercial_productos] b 
    ON a.[codProd] = b.[codProd]
WHERE a.mes BETWEEN '2501' AND '2604' and a.cstatus !='A' AND a.ctipo_vta='0003' and ccod_ruta!= '0000'
GROUP BY
      a.ccod_cli
	 ,a.[mes]
    ,a.[ccod_ruta]
	,b.[categoria]
	,b.linea
	,b.sublinea
	,b.producto
	,a.[fecha]
ORDER BY 
	a.[fecha] DESC