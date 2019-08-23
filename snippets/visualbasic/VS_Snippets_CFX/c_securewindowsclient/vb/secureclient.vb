﻿'<snippet0>
Imports System.Collections.Generic
Imports System.ServiceModel



Public Class Program
    
    Shared Sub Main() 
        '<snippet1>
        Dim b As New WSHttpBinding(SecurityMode.Message)
        b.Security.Message.ClientCredentialType = MessageCredentialType.Windows
        
        Dim ea As New EndpointAddress("net.tcp://machinename:8036/endpoint")
        Dim cc As New CalculatorClient(b, ea)
        cc.Open()
 
        ' Alternatively, use a binding name from a configuration file generated by the
        ' SvcUtil.exe tool to create the client. Omit the binding and endpoint address 
        ' because that information is provided by the configuration file.
        ' CalculatorClass cc = new CalculatorClient("ICalculator_Binding")
        '</snippet1>
    End Sub 
End Class 'Program


<System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0"), System.ServiceModel.ServiceContractAttribute([Namespace] := "http://Microsoft.ServiceModel.Samples", ConfigurationName := "ICalculator")>  _
Public Interface ICalculator
    
    <System.ServiceModel.OperationContractAttribute(Action := "http://Microsoft.ServiceModel.Samples/ICalculator/Add", ReplyAction := "http://Microsoft.ServiceModel.Samples/ICalculator/AddResponse")>  _
    Function Add(ByVal n1 As Double, ByVal n2 As Double) As Double 
    
    <System.ServiceModel.OperationContractAttribute(Action := "http://Microsoft.ServiceModel.Samples/ICalculator/Subtract", ReplyAction := "http://Microsoft.ServiceModel.Samples/ICalculator/SubtractResponse")>  _
    Function Subtract(ByVal n1 As Double, ByVal n2 As Double) As Double 
    
    <System.ServiceModel.OperationContractAttribute(Action := "http://Microsoft.ServiceModel.Samples/ICalculator/Multiply", ReplyAction := "http://Microsoft.ServiceModel.Samples/ICalculator/MultiplyResponse")>  _
    Function Multiply(ByVal n1 As Double, ByVal n2 As Double) As Double 
    
    <System.ServiceModel.OperationContractAttribute(Action := "http://Microsoft.ServiceModel.Samples/ICalculator/Divide", ReplyAction := "http://Microsoft.ServiceModel.Samples/ICalculator/DivideResponse")>  _
    Function Divide(ByVal n1 As Double, ByVal n2 As Double) As Double 
End Interface 'ICalculator

<System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")>  _
Public Interface ICalculatorChannel
    : Inherits ICalculator, System.ServiceModel.IClientChannel
End Interface 'ICalculatorChannel

<System.Diagnostics.DebuggerStepThroughAttribute(), System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")>  _
Public Class CalculatorClient
    Inherits System.ServiceModel.ClientBase(Of ICalculator)
    Implements ICalculator
     
    Public Sub New() 
     '
    End Sub 'New
    
    
    Public Sub New(ByVal endpointConfigurationName As String) 
        MyBase.New(endpointConfigurationName)
    
    End Sub 'New
    
    
    Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As String) 
        MyBase.New(endpointConfigurationName, remoteAddress)
    
    End Sub 'New
    
    
    Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As System.ServiceModel.EndpointAddress) 
        MyBase.New(endpointConfigurationName, remoteAddress)
    
    End Sub 'New
    
    
    Public Sub New(ByVal binding As System.ServiceModel.Channels.Binding, ByVal remoteAddress As System.ServiceModel.EndpointAddress) 
        MyBase.New(binding, remoteAddress)
    
    End Sub 'New
    
    
    Public Function Add(ByVal n1 As Double, ByVal n2 As Double) As Double Implements ICalculator.Add
        Return MyBase.Channel.Add(n1, n2)
    
    End Function 'Add
    
    
    Public Function Subtract(ByVal n1 As Double, ByVal n2 As Double) As Double Implements ICalculator.Subtract
        Return MyBase.Channel.Subtract(n1, n2)
    
    End Function 'Subtract
    
    
    Public Function Multiply(ByVal n1 As Double, ByVal n2 As Double) As Double Implements ICalculator.Multiply
        Return MyBase.Channel.Multiply(n1, n2)
    
    End Function 'Multiply
    
    
    Public Function Divide(ByVal n1 As Double, ByVal n2 As Double) As Double Implements ICalculator.Divide
        Return MyBase.Channel.Divide(n1, n2)
    
    End Function 'Divide
End Class 'CalculatorClient
'</snippet0>