using Microsoft.AspNetCore.Mvc;
using Microsoft.AnalysisServices.AdomdClient;
using System.Collections.Generic;

namespace YourNamespace.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TABULARController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public TABULARController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

       
        ////////////////////////////////////////////////////// Consulta #1 ////////////////////////////////////////////////////
        /************************************* Ventas de productos en fin de semana **************************************************************/
        [HttpGet("Ventas de productos en fin de semana")]
        public IActionResult VentasProductosFinSemana()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT \r\n    NON EMPTY { [Measures].[Cantidad ventas] } ON COLUMNS, \r\n    NON EMPTY { \r\n        ([TAB_SGV ViewProductos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * \r\n         [TAB_SGV ViewDate].[Fin de semana].[true]) \r\n    } \r\n    DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS \r\nFROM [Model] \r\nCELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #2 ////////////////////////////////////////////////////
        /************************************* Top ingresos por cliente **************************************************************/

        [HttpGet("Top ingresos por cliente")]
        public IActionResult TopIngresosClientes()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT \r\n    NON EMPTY { [Measures].[Ventas totales] } ON COLUMNS, \r\n    NON EMPTY { \r\n        HEAD(\r\n            ORDER(\r\n                [TAB_SGV ViewClientes].[Nombre Cliente].[Nombre Cliente].ALLMEMBERS, \r\n                [Measures].[Ventas totales], \r\n                DESC\r\n            ), \r\n            20\r\n        ) \r\n    } \r\n    DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS \r\nFROM [Model] \r\nCELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #3 ////////////////////////////////////////////////////
        /************************************* Venta de productos por proveedor **************************************************************/

        [HttpGet("Venta de productos por proveedor")]
        public IActionResult VentaProductosProveedor()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Cantidad ventas] } ON COLUMNS, NON EMPTY { ([TAB_SGV ViewProductos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [TAB_SGV ViewProductos].[Nombre Proveedor].[Nombre Proveedor].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Model] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        ////////////////////////////////////////////////////// Consulta #4 ////////////////////////////////////////////////////
        /************************************* Productos menos vendidos **************************************************************/

        [HttpGet("Productos menos vendidos")]
        public IActionResult ProductosMenosVendidos()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT \r\n    NON EMPTY { [Measures].[Cantidad ventas] } ON COLUMNS, \r\n    NON EMPTY { \r\n        ORDER(\r\n            [TAB_SGV ViewProductos].[Nombre Producto].[Nombre Producto].ALLMEMBERS, \r\n            [Measures].[Cantidad ventas], \r\n            ASC\r\n        ) \r\n    } \r\n    DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS \r\nFROM [Model] \r\nCELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #5 ////////////////////////////////////////////////////
        /************************************* Ventas semanales **************************************************************/

        [HttpGet("Ventas semanales")]
        public IActionResult VentasMensuales()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Ventas totales] } ON COLUMNS, NON EMPTY { ([TAB_SGV ViewDate].[Anio].[Anio].ALLMEMBERS * [TAB_SGV ViewDate].[Nombre del mes].[Nombre del mes].ALLMEMBERS * [TAB_SGV ViewDate].[Semana del mes].[Semana del mes].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Model] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        ////////////////////////////////////////////////////// Consulta #6 ////////////////////////////////////////////////////
        /************************************* Ingresos totales mensuales **************************************************************/


        [HttpGet("Ingresos totales mensuales")]
        public IActionResult IngresosSemanales()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Ventas totales] } ON COLUMNS, NON EMPTY { ([TAB_SGV ViewDate].[Nombre del mes].[Nombre del mes].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Model] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #7 ////////////////////////////////////////////////////
        /************************************* Promedio de ventas diarias **************************************************************/


        [HttpGet("Promedio de ventas diarias")]
        public IActionResult PromedioVentasDiarias()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "WITH \r\nMEMBER [Measures].[Promedio ventas diarias] AS \r\n    AVG(\r\n        EXISTING ([TAB_SGV ViewDate].[Nombre del dia].[Nombre del dia].ALLMEMBERS), \r\n        [Measures].[Ventas totales]\r\n    )\r\n\r\nSELECT \r\n    NON EMPTY { [Measures].[Promedio ventas diarias] } ON COLUMNS, \r\n    NON EMPTY { \r\n        [TAB_SGV ViewDate].[Nombre del mes].[Nombre del mes].ALLMEMBERS \r\n    } \r\n    DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS \r\nFROM [Model] \r\nCELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #8 ////////////////////////////////////////////////////
        /************************************* Ventas de productos por categoria **************************************************************/


        [HttpGet("Ventas de productos por categoria")]
        public IActionResult VentasProductosCategoria()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Ventas totales] } ON COLUMNS, NON EMPTY { ([TAB_SGV ViewProductos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [TAB_SGV ViewProductos].[Categoria Producto].[Categoria Producto].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Model] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        ////////////////////////////////////////////////////// Consulta #9 ////////////////////////////////////////////////////
        /************************************* Clientes con mayor cantidad de ventas **************************************************************/


        [HttpGet("Clientes con mayor cantidad de ventas")]
        public IActionResult ClientesMayorCantidadVenta()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT \r\n    NON EMPTY { [Measures].[Cantidad ventas] } ON COLUMNS, \r\n    NON EMPTY { \r\n        HEAD(\r\n            ORDER(\r\n                [TAB_SGV ViewClientes].[Nombre Cliente].[Nombre Cliente].ALLMEMBERS, \r\n                [Measures].[Cantidad ventas], \r\n                DESC\r\n            ), \r\n            20\r\n        ) \r\n    } \r\n    DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS \r\nFROM [Model] \r\nCELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        ////////////////////////////////////////////////////// Consulta #10 ////////////////////////////////////////////////////
        /************************************* Venta de productos por fecha **************************************************************/


        [HttpGet("Venta de productos por fecha")]
        public IActionResult VentaProductosFecha()
        {
            string connectionString = _configuration.GetConnectionString("TabularConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Ventas totales] } ON COLUMNS, NON EMPTY { ([TAB_SGV ViewProductos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [TAB_SGV ViewDate].[Nombre del mes].[Nombre del mes].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Model] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        private List<Dictionary<string, object>> TransformToJSON(CellSet result)
        {
            var jsonData = new List<Dictionary<string, object>>();
            int cellIndex = 0;

            foreach (var rowPosition in result.Axes[1].Positions)
            {
                var dataPoint = new Dictionary<string, object>();

                for (int i = 0; i < rowPosition.Members.Count; i++)
                {
                    var dimensionName = result.Axes[1].Set.Hierarchies[i].Name;
                    dataPoint[dimensionName] = rowPosition.Members[i].Caption;
                }

                for (int colIndex = 0; colIndex < result.Axes[0].Positions.Count; colIndex++)
                {
                    var measureName = result.Axes[0].Positions[colIndex].Members[0].Caption;
                    var cellValue = result.Cells[cellIndex].Value;

                    dataPoint[measureName] = cellValue;
                    cellIndex++;
                }

                jsonData.Add(dataPoint);
            }
            return jsonData;
        }
    }
}
