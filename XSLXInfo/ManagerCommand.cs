using System;
using System.Windows.Input;

namespace XLSXInfo
{
	public class ManagerCommand : ICommand
	{
		private Action action;
		private Action<object> porcWithParam;
		private Func<bool> canExecute;

		public event EventHandler CanExecuteChanged
		{
			add
			{
				CommandManager.RequerySuggested += value;
			}
			remove
			{
				CommandManager.RequerySuggested -= value;
			}
		}

		public ManagerCommand(Action action, Func<bool> canAction = null)
		{
			this.action = action;
			if (canAction == null)
			{
				this.canExecute = () => true;
			}
			else
			{
				this.canExecute = canAction;
			}
		}

		public ManagerCommand(Action<object> porcWithParam, Func<bool> canAction = null)
		{
			this.porcWithParam = porcWithParam;
			if (canAction == null)
			{
				this.canExecute = () => true;
			}
			else
			{
				this.canExecute = canAction;
			}
		}

		public bool CanExecute(object parameter)
		{
			return canExecute();
		}

		public void Execute(object parameter)
		{
			if (action != null)
			{
				action();
				return;
			}

			if (porcWithParam != null)
			{
				porcWithParam(parameter);
			}

		}
	}
}