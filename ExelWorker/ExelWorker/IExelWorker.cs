using System.Collections.Generic;
using System.IO;

namespace ExelWorker.ExelWorker
{
    public interface IExelWorker
    {
        List<TModel> GetModelFromExel<TModel>(FileStream fileStream)
            where TModel : class;
    }
}
