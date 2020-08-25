using System.Collections.Generic;
using System.IO;

namespace ExelWorker.ExelWorker
{
    public interface IExelWorkerService
    {
        List<TModel> GetModelFromExel<TModel>(Stream fileStream)
            where TModel : class;
    }
}
