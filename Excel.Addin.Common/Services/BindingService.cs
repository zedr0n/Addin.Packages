using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;
using ExcelInterfaces;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;


namespace CommonAddin
{
    public static class PropExtensions
    {
        // returns property setter:
        public static Action<TObject, TProperty> GetPropSetter<TObject, TProperty>(string propertyName)
        {
            ParameterExpression paramExpression = Expression.Parameter(typeof(TObject));

            ParameterExpression paramExpression2 = Expression.Parameter(typeof(TProperty), propertyName);

            MemberExpression propertyGetterExpression = Expression.Property(paramExpression, propertyName);

            Action<TObject, TProperty> result = Expression.Lambda<Action<TObject, TProperty>>
            (
                Expression.Assign(propertyGetterExpression, paramExpression2), paramExpression, paramExpression2
            ).Compile();

            return result;
        }
    }

    public class BindingService : IBindingService
    {
        private readonly Dictionary<string,Binding> _bindings = new Dictionary<string, Binding>();
        private readonly Dictionary<string, string> _cellHandles = new Dictionary<string, string>();

        public IAddressService AddressService { get; set; }

        public BindingService(IAddressService addressService)
        {
            AddressService = addressService;
        }

        public TProperty AddBinding<T,TProperty>(T obj, string propertyName) where T : class
        {
            var lambda = BindExtensions.GetPropertyEx<T, TProperty>(obj, propertyName);
            AddBinding(obj,lambda);
            return lambda.Compile()(obj);
        }

        public TProperty AddBinding<T,TProperty>(T obj, Expression<Func<T, TProperty>> memberLambda) where T : class
        {
            var cell = AddressService.GetAddress();
            var formula = GetFormula();
            _bindings[cell] = new Binding<T, TProperty>(cell, obj, memberLambda,formula);
            return memberLambda.Compile()(obj);
        }

        public void OnValueChanged(Range range)
        {
            var cell = $"[{range?.Application.ActiveWorkbook.Name}]{range?.Worksheet.Name}{range?.Address}";
            var value = range?.Value;

            if (_bindings.All(x => x.Key != cell))
                return;
            var binding = _bindings[cell];
            if (value == null || (string) value == "")
                _bindings[cell] = null;
            else
            {
                binding?.Set(value);
                // restore the formula
                if ((string) range.FormulaR1C1 != binding.Formula)
                    range.FormulaR1C1 = binding.Formula;
            }

            // reset the handle association on change
            //_cellHandles[cell] = null;
        }

        public string GetFormula()
        {
            var reference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (reference == null)
                return null;

            var xlCell = (Range)ReferenceToRange(reference);
            return (string) xlCell.FormulaR1C1;
        }

        private static object ReferenceToRange(ExcelReference xlref)
        {
            var app = ExcelDnaUtil.Application;
            var refText = XlCall.Excel(XlCall.xlfReftext, xlref, true);
            var range = app.GetType().InvokeMember("Range",
                BindingFlags.Public | BindingFlags.GetProperty,
                null, app, new object[] { refText });
            return range;
        }

        /// <summary>
        /// Syncs the property values for the object bound to changed cell
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="target"></param>
        public void OnSheetChange(object sheet, object target)
        {
            var range = target as Microsoft.Office.Interop.Excel.Range;
            OnValueChanged(range);

            // What to do with cut and paste???
            // the address should update
        }
    }

    public class Binding
    {
        public virtual void Set(object value) { }
        public string Formula { get; set; }
    }

    public class Binding<T,TProperty> : Binding where T : class
    {
        // cell formula
        // cell reference
        private string _cell;
        // associated object 
        private T _object;
        // function which is associated with the object action cell
        private readonly Action<TProperty> _property;

        public override void Set(object value)
        {
            _property((TProperty) value);
        }

        public Binding(string cell, T obj, Expression<Func<T, TProperty>> memberLamda, string formula = "")
        {
            _cell = cell;
            _object = obj;
            Formula = formula;

            var memberSelectorExpression = memberLamda.Body as MemberExpression;
            if (memberSelectorExpression != null)
            {
                var property = memberSelectorExpression.Member as PropertyInfo;
                if (property != null)
                {
                    //    _property = value => property.SetValue(obj, value, null);
                    var setter = PropExtensions.GetPropSetter<T, TProperty>(property.Name);
                    _property = x => setter(obj,x);
                }   
            }
        }
    }
}