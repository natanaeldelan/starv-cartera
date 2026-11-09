using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.Sqlite;
using Dapper;
using System.Net;

var builder = WebApplication.CreateBuilder(args);

// Configurar el puerto dinámico que exige la nube (Render, Railway, etc.)
var port = Environment.GetEnvironmentVariable("PORT") ?? "5000";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

var app = builder.Build();

// Base de datos SQLite local (en la nube se mantiene segura en el contenedor)
string connectionString = "Data Source=cartera_startv.db";

using (var connection = new SqliteConnection(connectionString))
{
    connection.Open();
    string createTableQuery = @"
        CREATE TABLE IF NOT EXISTS SeguimientoClientes (
            Contrato TEXT PRIMARY KEY,
            Cliente TEXT,
            Plaza TEXT,
            FechaContratacion TEXT,
            Telefono TEXT,
            MesesPagados INTEGER,
            Tecnico TEXT,
            Estatus TEXT DEFAULT 'Pendiente',
            Observaciones TEXT,
            Liquidado INTEGER DEFAULT 0
        );";
    connection.Execute(createTableQuery);
}

// 1. Pantalla Principal
app.MapGet("/", async context =>
{
    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.WriteAsync(@"
        <!DOCTYPE html>
        <html>
        <head>
            <title>Sistema de Cartera StarTV</title>
            <link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'>
            <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css'>
        </head>
        <body class='bg-light p-4'>
            <div class='container bg-white p-4 rounded shadow-sm'>
                <h2 class='text-primary mb-2'><i class='bi bi-cloud-check-fill'></i> Control y Seguimiento de Cartera - StarTV (Nube)</h2>
                <p class='text-muted'>Sistema centralizado accesible desde cualquier ubicación.</p>
                
                <div class='d-flex gap-3 mb-4'>
                    <a href='/dashboard' class='btn btn-dark'><i class='bi bi-kanban'></i> Ver Tablero de Seguimiento Actual</a>
                </div>

                <form action='/upload' method='post' enctype='multipart/form-data'>
                    <div class='mb-3'>
                        <label class='form-label fw-bold'>Subir o Actualizar Concentrado de Clientes (.xlsx):</label>
                        <input type='file' name='archivoExcel' class='form-control' accept='.xlsx' required />
                    </div>
                    
                    <div class='mb-4 p-3 bg-light rounded border'>
                        <label class='form-label fw-bold text-dark'>Filtro de Meses Pagados para Cargar a BD:</label>
                        <div class='row row-cols-auto g-2 mt-1'>");

    for (int i = 0; i <= 12; i++)
    {
        string isChecked = (i <= 3) ? "checked" : "";
        await context.Response.WriteAsync($@"
                            <div class='col-auto me-2'>
                                <div class='form-check'>
                                    <input class='form-check-input' type='checkbox' name='mesesFiltro' value='{i}' id='mes_{i}' {isChecked}>
                                    <label class='form-check-label fw-semibold' for='mes_{i}'>{i} {(i == 1 ? "mes" : "meses")}</label>
                                </div>
                            </div>");
    }

    await context.Response.WriteAsync(@"
                        </div>
                    </div>

                    <button type='submit' class='btn btn-primary btn-lg'><i class='bi bi-cloud-arrow-up-fill'></i> Procesar y Actualizar Base de Datos</button>
                </form>
            </div>
        </body>
        </html>");
});

// 2. Procesar Excel y Guardar/Actualizar en SQLite
app.MapPost("/upload", async (HttpContext context) =>
{
    var form = await context.Request.ReadFormAsync();
    var file = form.Files["archivoExcel"];

    if (file == null || file.Length == 0)
    {
        await context.Response.WriteAsync("Por favor selecciona un archivo válido.");
        return;
    }

    var mesesSeleccionados = form["mesesFiltro"]
        .Select(val => int.TryParse(val, out int m) ? m : -1)
        .Where(m => m >= 0)
        .ToHashSet();

    using (var stream = file.OpenReadStream())
    using (var workbook = new XLWorkbook(stream))
    {
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RowsUsed().Skip(2);

        using var connection = new SqliteConnection(connectionString);
        connection.Open();

        foreach (var row in rows)
        {
            string plaza = row.Cell(6).GetString().Trim();
            string contrato = row.Cell(9).GetString().Trim();
            string cliente = row.Cell(7).GetString().Trim();
            string fechaContratacion = row.Cell(30).GetString().Trim();
            string telefono = row.Cell(51).GetString().Trim();
            string celular = row.Cell(53).GetString().Trim();
            string tecnico = row.Cell(59).GetString().Trim();
            
            if (string.IsNullOrEmpty(tecnico)) tecnico = "TÉCNICO NO ASIGNADO / GENERAL";
            string contactoTel = !string.IsNullOrEmpty(celular) && celular != "0" ? celular : (!string.IsNullOrEmpty(telefono) && telefono != "0" ? telefono : "S/N");

            int mesesPagados = 0;
            if (int.TryParse(row.Cell(80).GetString().Trim(), out int parsedPagos))
            {
                mesesPagados = parsedPagos;
            }

            if (mesesSeleccionados.Contains(mesesPagados) && !string.IsNullOrEmpty(contrato))
            {
                string upsertQuery = @"
                    INSERT INTO SeguimientoClientes (Contrato, Cliente, Plaza, FechaContratacion, Telefono, MesesPagados, Tecnico)
                    VALUES (@Contrato, @Cliente, @Plaza, @FechaContratacion, @Telefono, @MesesPagados, @Tecnico)
                    ON CONFLICT(Contrato) DO UPDATE SET 
                        MesesPagados = @MesesPagados,
                        Telefono = @Telefono,
                        Tecnico = @Tecnico;
                ";
                connection.Execute(upsertQuery, new { Contrato = contrato, Cliente = cliente, Plaza = plaza, FechaContratacion = fechaContratacion, Telefono = contactoTel, MesesPagados = mesesPagados, Tecnico = tecnico });
            }
        }
    }

    context.Response.Redirect("/dashboard");
});

// 3. Tablero de Seguimiento
app.MapGet("/dashboard", async context =>
{
    using var connection = new SqliteConnection(connectionString);
    var cuentas = connection.Query("SELECT * FROM SeguimientoClientes WHERE Liquidado = 0 ORDER BY Tecnico, MesesPagados").ToList();

    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.WriteAsync("<!DOCTYPE html><html><head>");
    await context.Response.WriteAsync("<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'>");
    await context.Response.WriteAsync("<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css'>");
    await context.Response.WriteAsync("</head><body class='bg-light p-4'><div class='container-fluid bg-white p-4 rounded shadow-sm'>");
    
    await context.Response.WriteAsync("<div class='d-flex justify-content-between align-items-center mb-4'>");
    await context.Response.WriteAsync("<div><h2 class='text-primary mb-0'>Tablero de Seguimiento de Cartera</h2><small class='text-muted'>Acceso global para el equipo</small></div>");
    await context.Response.WriteAsync("<div><a href='/' class='btn btn-secondary me-2'><i class='bi bi-cloud-arrow-up'></i> Subir Nuevo Excel</a><a href='/liquidados' class='btn btn-success'><i class='bi bi-check-all'></i> Cuentas Liquidadas</a></div>");
    await context.Response.WriteAsync("</div>");

    var grupos = cuentas.GroupBy(x => x.Tecnico);

    foreach (var grupo in grupos)
    {
        await context.Response.WriteAsync($"<div class='card mb-4 border-primary'><div class='card-header bg-primary text-white'>");
        await context.Response.WriteAsync($"<h5 class='mb-0'><i class='bi bi-person-badge'></i> Técnico: {grupo.Key} <span class='badge bg-light text-dark ms-2'>{grupo.Count()} cuentas</span></h5>");
        await context.Response.WriteAsync($"</div><div class='card-body table-responsive'><table class='table table-striped table-hover table-sm align-middle mb-0'>");
        await context.Response.WriteAsync("<thead><tr><th>Contrato</th><th>Cliente</th><th>Teléfono</th><th>Plaza</th><th>Fecha Inst.</th><th>Pagos</th><th>Observaciones / Avances</th><th>Acciones</th></tr></thead><tbody>");

        foreach (var item in grupo)
        {
            string mensajeWp = $"Hola, te comparto el caso de StarTV Contrato {item.Contrato} ({item.Cliente}), tel: {item.Telefono}, plaza {item.Plaza}. Registra {item.MesesPagados} meses pagados. Observación: {item.Observaciones}";
            string urlWp = $"https://wa.me/?text={WebUtility.UrlEncode(mensajeWp)}";

            await context.Response.WriteAsync($"<tr>");
            await context.Response.WriteAsync($"<td><strong>{item.Contrato}</strong></td>");
            await context.Response.WriteAsync($"<td>{item.Cliente}</td>");
            await context.Response.WriteAsync($"<td><span class='badge bg-info text-dark'>{item.Telefono}</span></td>");
            await context.Response.WriteAsync($"<td>{item.Plaza}</td>");
            await context.Response.WriteAsync($"<td>{item.FechaContratacion}</td>");
            await context.Response.WriteAsync($"<td><span class='badge bg-warning text-dark'>{item.MesesPagados} pagos</span></td>");
            
            await context.Response.WriteAsync($"<td>");
            await context.Response.WriteAsync($"<form action='/update-obs' method='post' class='d-flex gap-1'>");
            await context.Response.WriteAsync($"<input type='hidden' name='contrato' value='{item.Contrato}' />");
            await context.Response.WriteAsync($"<input type='text' name='observaciones' value='{item.Observaciones}' class='form-control form-control-sm' placeholder='Escribir avance...' />");
            await context.Response.WriteAsync($"<button type='submit' class='btn btn-outline-secondary btn-sm' title='Guardar'><i class='bi bi-save'></i></button>");
            await context.Response.WriteAsync($"</form>");
            await context.Response.WriteAsync($"</td>");

            await context.Response.WriteAsync($"<td class='text-nowrap'>");
            await context.Response.WriteAsync($"<a href='{urlWp}' target='_blank' class='btn btn-success btn-sm me-1' title='Enviar WhatsApp'><i class='bi bi-whatsapp'></i></a>");
            await context.Response.WriteAsync($"<a href='/liquidar?contrato={item.Contrato}' class='btn btn-outline-success btn-sm me-1' title='Liquidar Cuenta'><i class='bi bi-check-lg'></i> Liquidar</a>");
            await context.Response.WriteAsync($"<a href='/delete?contrato={item.Contrato}' class='btn btn-outline-danger btn-sm' onclick=\"return confirm('¿Borrar registro?');\" title='Eliminar'><i class='bi bi-trash'></i></a>");
            await context.Response.WriteAsync($"</td>");
            await context.Response.WriteAsync($"</tr>");
        }
        await context.Response.WriteAsync("</tbody></table></div></div>");
    }

    await context.Response.WriteAsync("</div></body></html>");
});

// 4. Actualizar Observaciones
app.MapPost("/update-obs", async (HttpContext context) =>
{
    var form = await context.Request.ReadFormAsync();
    string contrato = form["contrato"];
    string observaciones = form["observaciones"];

    using var connection = new SqliteConnection(connectionString);
    connection.Execute("UPDATE SeguimientoClientes SET Observaciones = @Obs WHERE Contrato = @Contrato", new { Obs = observaciones, Contrato = contrato });

    context.Response.Redirect("/dashboard");
});

// 5. Liquidar Cuenta
app.MapGet("/liquidar", async context =>
{
    string contrato = context.Request.Query["contrato"];
    if (!string.IsNullOrEmpty(contrato))
    {
        using var connection = new SqliteConnection(connectionString);
        connection.Execute("UPDATE SeguimientoClientes SET Liquidado = 1 WHERE Contrato = @Contrato", new { Contrato = contrato });
    }
    context.Response.Redirect("/dashboard");
});

// 6. Eliminar Registro
app.MapGet("/delete", async context =>
{
    string contrato = context.Request.Query["contrato"];
    if (!string.IsNullOrEmpty(contrato))
    {
        using var connection = new SqliteConnection(connectionString);
        connection.Execute("DELETE FROM SeguimientoClientes WHERE Contrato = @Contrato", new { Contrato = contrato });
    }
    context.Response.Redirect("/dashboard");
});

// 7. Ver Cuentas Liquidadas
app.MapGet("/liquidados", async context =>
{
    using var connection = new SqliteConnection(connectionString);
    var cuentas = connection.Query("SELECT * FROM SeguimientoClientes WHERE Liquidado = 1").ToList();

    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.WriteAsync("<!DOCTYPE html><html><head><link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'></head><body class='bg-light p-4'><div class='container bg-white p-4 rounded shadow-sm'>");
    await context.Response.WriteAsync("<h2 class='text-success mb-3'>Cuentas Liquidadas / Resueltas</h2><a href='/dashboard' class='btn btn-secondary mb-3'>Regresar al Tablero</a>");
    
    await context.Response.WriteAsync("<table class='table table-bordered table-sm'><thead><tr><th>Contrato</th><th>Cliente</th><th>Plaza</th><th>Técnico</th><th>Observaciones</th></tr></thead><tbody>");
    foreach (var item in cuentas)
    {
        await context.Response.WriteAsync($"<tr><td>{item.Contrato}</td><td>{item.Cliente}</td><td>{item.Plaza}</td><td>{item.Tecnico}</td><td>{item.Observaciones}</td></tr>");
    }
    await context.Response.WriteAsync("</tbody></table></div></div></body></html>");
});

app.Run();