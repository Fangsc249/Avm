using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
 
// 暂未使用
// 2025-3-27
namespace AutomationVM.Core
{
    public interface IResettable
    {
        void Reset();
    }
    public class InstructionPool<T> where T : Instruction, new()
    {
        private readonly ConcurrentBag<T> _pool = new ConcurrentBag<T>();

        public T Rent()
        {
            return _pool.TryTake(out T item) ? item : new T();
        }
        public void Return(T item)
        {
            if (item is IResettable resettable) resettable.Reset();
            _pool.Add(item);
        }
    }
}
