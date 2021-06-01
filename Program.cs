using System;

namespace aula04
{
    class Program
    {
        static void Main(string[] args)
        {
            //Calculadora simples

            Inicio:
            
            Console.Clear();
            
            Console.Write("Digite um valor: ");
            double num1 = double.Parse(Console.ReadLine());

            Console.Write("Digite outro valor: ");
            double num2 = double.Parse(Console.ReadLine());

            Console.Write("Operação que deseja realizar: + - x /");
            char op = char.Parse(Console.ReadLine());

            double resultado;

            switch (op)

            {
                default: 
                    Console.WriteLine("Erro na operação");
                 break;

                case '+': 
                    resultado = num1 + num2;
                    Console.WriteLine("O resultado é: " + resultado);
                break;

                case '-':
                    resultado = num1 - num2;
                    Console.WriteLine("O resultado é: " + resultado);
                break;

                case 'x':
                    resultado = num1 * num2;
                    Console.WriteLine("O resultado é: " + resultado);
                break;

                case '/':
                    
                    if (num2 == 0)
                        {
                        Console.WriteLine("O número não pode ser dividido por 0");
                        }
                        else 
                        {
                            resultado = num1 / num2;
                         Console.WriteLine("O resultado é: " + resultado);
                        }
                                         
                break;
            }

            Console.WriteLine("Deseja continuar Calculando? [y/n]");
            string opcao = Console.ReadLine();
                if (opcao == "y" || opcao == "Y")
                {    
                goto Inicio;
                }
             Console.ReadKey();









        }
    }
}
