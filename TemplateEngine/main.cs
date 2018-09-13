using System.Activities;
using System.ComponentModel;
using System.Data;
using HandlebarsDotNet;
using System.IO;
using System.Collections.Generic;
using System;

namespace TemplateEngine
{
    public class TemplateEngine : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> TemplateFilename { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<DataTable> Parameters { get; set; }

        [Category("Output")]
        public OutArgument<string> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //Get template and parameters
            string filename = TemplateFilename.Get(context);
            DataTable parameters = Parameters.Get(context);

            //RegisterHelper - Tracking current filename & include depth
            string currentFile = filename;
            const int maxDepth = 30;
            int curDepth = 0;

            //Convert DataTable to useable notation
            Dictionary<string, object> data = DataTableToDict(parameters);

            //Register include helper function
            Handlebars.RegisterHelper("include", (writer, include_context, include_params) =>
            {
                if (include_params.Length != 1)
                {
                    throw new ArgumentException("In template file (" + currentFile + ") Use of include directive requires exactly 1 argument - filename.");
                }
                string subTemplateText = LoadTemplateFile(include_params[0].ToString(), currentFile);


                string previousFilename = currentFile;
                currentFile = include_params[0].ToString();

                curDepth++;
                if (curDepth > maxDepth)
                {
                    throw new ArgumentException("Maximum include depth (" + maxDepth.ToString() + ") reached in file " + previousFilename);
                }

                string subResult = RenderTemplate(subTemplateText, include_context, currentFile);
                curDepth--;

                currentFile = previousFilename;

                writer.WriteSafeString(subResult);
            });

            //Run handlebars on template + parameters
            string templateText = LoadTemplateFile(filename, "In UIPath Studio");
            string result = RenderTemplate(templateText, data, filename);



            //Output the result
            Result.Set(context, result);

        }

        //Load a template file from file
        static string LoadTemplateFile(string filename, string currentFile)
        {
            string templateText;
            try
            {
                templateText = File.ReadAllText(filename);
            }
            catch (FileNotFoundException e)
            {
                throw new FileNotFoundException("In template file (" + currentFile + ") could not find file included : " + e.FileName);
            }
            catch (ArgumentException e)
            {
                throw new ArgumentException("In template file (" + currentFile + ") Include directive error: " + e.Message);
            }
            catch
            {
                throw new SystemException("Could not load template file " + filename);
            }
            return templateText;
        }

        //Render the template using the given parameters (context)
        //Filename is for error tracing
        static string RenderTemplate(string templateText, object context, string filename)
        {
            string result;
            try
            {
                var template = Handlebars.Compile(templateText);
                result = template(context);
            }
            catch (HandlebarsParserException e)
            {

                throw new ArgumentException("In template file (" + filename + ") " + e.Message);

            }
            catch (HandlebarsCompilerException e)
            {

                throw new ArgumentException("In template file (" + filename + ") " + e.Message);
            }
            catch
            {
                throw new SystemException("Could not render template " + filename);
            }

            return result;
        }

        //Convert Master DataTable to Dictionary
        static Dictionary<string, object> DataTableToDict(DataTable dtIn)
        {
            if (dtIn.Columns.Count != 2)
            {
                throw new ArgumentException("DataTable passed directly to Template Engine must contain exactly 2 columns");
            }

            if (dtIn.Columns[0].DataType != typeof(string))
            {
                throw new ArgumentException("DataTable passed directly to Template Engine must have first column set to Type: String");
            }

            Dictionary<string, object> dictOut = new Dictionary<string, object>();

            foreach (DataRow row in dtIn.Rows)
            {
                string key = row.Field<string>(0);
                object val = row.Field<object>(1);

                if (val.GetType() == typeof(DataTable))
                {
                    val = SubDataTableToList((DataTable)val);
                }

                dictOut.Add(key, val);
            }
            return dictOut;
        }

        //Convert Sub-DataTables to Lists of Dictionaries
        static List<Dictionary<string, object>> SubDataTableToList(DataTable dtIn)
        {
            List<Dictionary<string, object>> listOut = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;

            foreach (DataRow dr in dtIn.Rows)
            {
                row = new Dictionary<string, object>();
                foreach (DataColumn dc in dtIn.Columns)
                {
                    string key = dc.ColumnName;
                    object val = dr[dc];

                    if (val.GetType() == typeof(DataTable))
                    {
                        val = SubDataTableToList((DataTable)val);
                    }
                    row.Add(key, val);
                }
                listOut.Add(row);
            }
            return listOut;
        }
    }
}

