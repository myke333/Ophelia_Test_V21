﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Este código fue generado por una herramienta.
//     Versión de runtime:4.0.30319.42000
//
//     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
//     se vuelve a generar el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Ophelia_Test_V21.SmenwSPEnv {
    
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.ServiceContractAttribute(Namespace="http://kactus", ConfigurationName="SmenwSPEnv.IKGnSmenwWCF", SessionMode=System.ServiceModel.SessionMode.Required)]
    public interface IKGnSmenwWCF {
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_ApplyUpdates", ReplyAction="http://kactus/IKGnSmenwWCF/AS_ApplyUpdatesResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesResponse AS_ApplyUpdates(Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_ApplyUpdates", ReplyAction="http://kactus/IKGnSmenwWCF/AS_ApplyUpdatesResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesResponse> AS_ApplyUpdatesAsync(Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetRecords", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetRecordsResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsResponse AS_GetRecords(Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetRecords", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetRecordsResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsResponse> AS_GetRecordsAsync(Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_DataRequest", ReplyAction="http://kactus/IKGnSmenwWCF/AS_DataRequestResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        object AS_DataRequest(string ProviderName, object[] Data);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_DataRequest", ReplyAction="http://kactus/IKGnSmenwWCF/AS_DataRequestResponse")]
        System.Threading.Tasks.Task<object> AS_DataRequestAsync(string ProviderName, object[] Data);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetProviderNames", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetProviderNamesResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        object AS_GetProviderNames();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetProviderNames", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetProviderNamesResponse")]
        System.Threading.Tasks.Task<object> AS_GetProviderNamesAsync();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetParams", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetParamsResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsResponse AS_GetParams(Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_GetParams", ReplyAction="http://kactus/IKGnSmenwWCF/AS_GetParamsResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsResponse> AS_GetParamsAsync(Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_RowRequest", ReplyAction="http://kactus/IKGnSmenwWCF/AS_RowRequestResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestResponse AS_RowRequest(Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_RowRequest", ReplyAction="http://kactus/IKGnSmenwWCF/AS_RowRequestResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestResponse> AS_RowRequestAsync(Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_Execute", ReplyAction="http://kactus/IKGnSmenwWCF/AS_ExecuteResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteResponse AS_Execute(Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/AS_Execute", ReplyAction="http://kactus/IKGnSmenwWCF/AS_ExecuteResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteResponse> AS_ExecuteAsync(Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/ExecRemoteFunction", ReplyAction="http://kactus/IKGnSmenwWCF/ExecRemoteFunctionResponse")]
        [System.ServiceModel.ServiceKnownTypeAttribute(typeof(object[]))]
        Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionResponse ExecRemoteFunction(Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest request);
        
        // CODEGEN: Generando contrato de mensaje, ya que la operación tiene múltiples valores de devolución.
        [System.ServiceModel.OperationContractAttribute(Action="http://kactus/IKGnSmenwWCF/ExecRemoteFunction", ReplyAction="http://kactus/IKGnSmenwWCF/ExecRemoteFunctionResponse")]
        System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionResponse> ExecRemoteFunctionAsync(Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest request);
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_ApplyUpdates", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_ApplyUpdatesRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string ProviderName;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object Delta;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public int MaxErrors;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=3)]
        public object OwnerData;
        
        public AS_ApplyUpdatesRequest() {
        }
        
        public AS_ApplyUpdatesRequest(string ProviderName, object Delta, int MaxErrors, object OwnerData) {
            this.ProviderName = ProviderName;
            this.Delta = Delta;
            this.MaxErrors = MaxErrors;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_ApplyUpdatesResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_ApplyUpdatesResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public object AS_ApplyUpdatesResult;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public int ErrorCount;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public object OwnerData;
        
        public AS_ApplyUpdatesResponse() {
        }
        
        public AS_ApplyUpdatesResponse(object AS_ApplyUpdatesResult, int ErrorCount, object OwnerData) {
            this.AS_ApplyUpdatesResult = AS_ApplyUpdatesResult;
            this.ErrorCount = ErrorCount;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_GetRecords", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_GetRecordsRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string ProviderName;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public int Count;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public int Options;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=3)]
        public string CommandText;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=4)]
        public object[] Params;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=5)]
        public object OwnerData;
        
        public AS_GetRecordsRequest() {
        }
        
        public AS_GetRecordsRequest(string ProviderName, int Count, int Options, string CommandText, object[] Params, object OwnerData) {
            this.ProviderName = ProviderName;
            this.Count = Count;
            this.Options = Options;
            this.CommandText = CommandText;
            this.Params = Params;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_GetRecordsResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_GetRecordsResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public object AS_GetRecordsResult;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public int RecsOut;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public object[] Params;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=3)]
        public object OwnerData;
        
        public AS_GetRecordsResponse() {
        }
        
        public AS_GetRecordsResponse(object AS_GetRecordsResult, int RecsOut, object[] Params, object OwnerData) {
            this.AS_GetRecordsResult = AS_GetRecordsResult;
            this.RecsOut = RecsOut;
            this.Params = Params;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_GetParams", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_GetParamsRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string ProviderName;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object OwnerData;
        
        public AS_GetParamsRequest() {
        }
        
        public AS_GetParamsRequest(string ProviderName, object OwnerData) {
            this.ProviderName = ProviderName;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_GetParamsResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_GetParamsResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public object AS_GetParamsResult;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object OwnerData;
        
        public AS_GetParamsResponse() {
        }
        
        public AS_GetParamsResponse(object AS_GetParamsResult, object OwnerData) {
            this.AS_GetParamsResult = AS_GetParamsResult;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_RowRequest", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_RowRequestRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string ProviderName;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object Row;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public int RequestType;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=3)]
        public object OwnerData;
        
        public AS_RowRequestRequest() {
        }
        
        public AS_RowRequestRequest(string ProviderName, object Row, int RequestType, object OwnerData) {
            this.ProviderName = ProviderName;
            this.Row = Row;
            this.RequestType = RequestType;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_RowRequestResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_RowRequestResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public object AS_RowRequestResult;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object OwnerData;
        
        public AS_RowRequestResponse() {
        }
        
        public AS_RowRequestResponse(object AS_RowRequestResult, object OwnerData) {
            this.AS_RowRequestResult = AS_RowRequestResult;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_Execute", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_ExecuteRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string ProviderName;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public string CommandText;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public object Params;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=3)]
        public object OwnerData;
        
        public AS_ExecuteRequest() {
        }
        
        public AS_ExecuteRequest(string ProviderName, string CommandText, object Params, object OwnerData) {
            this.ProviderName = ProviderName;
            this.CommandText = CommandText;
            this.Params = Params;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="AS_ExecuteResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class AS_ExecuteResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public object Params;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object OwnerData;
        
        public AS_ExecuteResponse() {
        }
        
        public AS_ExecuteResponse(object Params, object OwnerData) {
            this.Params = Params;
            this.OwnerData = OwnerData;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="ExecRemoteFunction", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class ExecRemoteFunctionRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public string pStMetodo;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object[] pDataIn;
        
        public ExecRemoteFunctionRequest() {
        }
        
        public ExecRemoteFunctionRequest(string pStMetodo, object[] pDataIn) {
            this.pStMetodo = pStMetodo;
            this.pDataIn = pDataIn;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="ExecRemoteFunctionResponse", WrapperNamespace="http://kactus", IsWrapped=true)]
    public partial class ExecRemoteFunctionResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=0)]
        public int ExecRemoteFunctionResult;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=1)]
        public object[] pDataOut;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://kactus", Order=2)]
        public string pStError;
        
        public ExecRemoteFunctionResponse() {
        }
        
        public ExecRemoteFunctionResponse(int ExecRemoteFunctionResult, object[] pDataOut, string pStError) {
            this.ExecRemoteFunctionResult = ExecRemoteFunctionResult;
            this.pDataOut = pDataOut;
            this.pStError = pStError;
        }
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public interface IKGnSmenwWCFChannel : Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF, System.ServiceModel.IClientChannel {
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public partial class KGnSmenwWCFClient : System.ServiceModel.ClientBase<Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF>, Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF {
        
        public KGnSmenwWCFClient() {
        }
        
        public KGnSmenwWCFClient(string endpointConfigurationName) : 
                base(endpointConfigurationName) {
        }
        
        public KGnSmenwWCFClient(string endpointConfigurationName, string remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public KGnSmenwWCFClient(string endpointConfigurationName, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public KGnSmenwWCFClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(binding, remoteAddress) {
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.AS_ApplyUpdates(Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest request) {
            return base.Channel.AS_ApplyUpdates(request);
        }
        
        public object AS_ApplyUpdates(string ProviderName, object Delta, int MaxErrors, ref object OwnerData, out int ErrorCount) {
            Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest();
            inValue.ProviderName = ProviderName;
            inValue.Delta = Delta;
            inValue.MaxErrors = MaxErrors;
            inValue.OwnerData = OwnerData;
            Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).AS_ApplyUpdates(inValue);
            ErrorCount = retVal.ErrorCount;
            OwnerData = retVal.OwnerData;
            return retVal.AS_ApplyUpdatesResult;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesResponse> AS_ApplyUpdatesAsync(Ophelia_Test_V21.SmenwSPEnv.AS_ApplyUpdatesRequest request) {
            return base.Channel.AS_ApplyUpdatesAsync(request);
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.AS_GetRecords(Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest request) {
            return base.Channel.AS_GetRecords(request);
        }
        
        public object AS_GetRecords(string ProviderName, int Count, int Options, string CommandText, ref object[] Params, ref object OwnerData, out int RecsOut) {
            Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest();
            inValue.ProviderName = ProviderName;
            inValue.Count = Count;
            inValue.Options = Options;
            inValue.CommandText = CommandText;
            inValue.Params = Params;
            inValue.OwnerData = OwnerData;
            Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).AS_GetRecords(inValue);
            RecsOut = retVal.RecsOut;
            Params = retVal.Params;
            OwnerData = retVal.OwnerData;
            return retVal.AS_GetRecordsResult;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsResponse> AS_GetRecordsAsync(Ophelia_Test_V21.SmenwSPEnv.AS_GetRecordsRequest request) {
            return base.Channel.AS_GetRecordsAsync(request);
        }
        
        public object AS_DataRequest(string ProviderName, object[] Data) {
            return base.Channel.AS_DataRequest(ProviderName, Data);
        }
        
        public System.Threading.Tasks.Task<object> AS_DataRequestAsync(string ProviderName, object[] Data) {
            return base.Channel.AS_DataRequestAsync(ProviderName, Data);
        }
        
        public object AS_GetProviderNames() {
            return base.Channel.AS_GetProviderNames();
        }
        
        public System.Threading.Tasks.Task<object> AS_GetProviderNamesAsync() {
            return base.Channel.AS_GetProviderNamesAsync();
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.AS_GetParams(Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest request) {
            return base.Channel.AS_GetParams(request);
        }
        
        public object AS_GetParams(string ProviderName, ref object OwnerData) {
            Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest();
            inValue.ProviderName = ProviderName;
            inValue.OwnerData = OwnerData;
            Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).AS_GetParams(inValue);
            OwnerData = retVal.OwnerData;
            return retVal.AS_GetParamsResult;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsResponse> AS_GetParamsAsync(Ophelia_Test_V21.SmenwSPEnv.AS_GetParamsRequest request) {
            return base.Channel.AS_GetParamsAsync(request);
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.AS_RowRequest(Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest request) {
            return base.Channel.AS_RowRequest(request);
        }
        
        public object AS_RowRequest(string ProviderName, object Row, int RequestType, ref object OwnerData) {
            Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest();
            inValue.ProviderName = ProviderName;
            inValue.Row = Row;
            inValue.RequestType = RequestType;
            inValue.OwnerData = OwnerData;
            Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).AS_RowRequest(inValue);
            OwnerData = retVal.OwnerData;
            return retVal.AS_RowRequestResult;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestResponse> AS_RowRequestAsync(Ophelia_Test_V21.SmenwSPEnv.AS_RowRequestRequest request) {
            return base.Channel.AS_RowRequestAsync(request);
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.AS_Execute(Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest request) {
            return base.Channel.AS_Execute(request);
        }
        
        public void AS_Execute(string ProviderName, string CommandText, ref object Params, ref object OwnerData) {
            Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest();
            inValue.ProviderName = ProviderName;
            inValue.CommandText = CommandText;
            inValue.Params = Params;
            inValue.OwnerData = OwnerData;
            Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).AS_Execute(inValue);
            Params = retVal.Params;
            OwnerData = retVal.OwnerData;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteResponse> AS_ExecuteAsync(Ophelia_Test_V21.SmenwSPEnv.AS_ExecuteRequest request) {
            return base.Channel.AS_ExecuteAsync(request);
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionResponse Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF.ExecRemoteFunction(Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest request) {
            return base.Channel.ExecRemoteFunction(request);
        }
        
        public int ExecRemoteFunction(string pStMetodo, object[] pDataIn, out object[] pDataOut, out string pStError) {
            Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest inValue = new Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest();
            inValue.pStMetodo = pStMetodo;
            inValue.pDataIn = pDataIn;
            Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionResponse retVal = ((Ophelia_Test_V21.SmenwSPEnv.IKGnSmenwWCF)(this)).ExecRemoteFunction(inValue);
            pDataOut = retVal.pDataOut;
            pStError = retVal.pStError;
            return retVal.ExecRemoteFunctionResult;
        }
        
        public System.Threading.Tasks.Task<Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionResponse> ExecRemoteFunctionAsync(Ophelia_Test_V21.SmenwSPEnv.ExecRemoteFunctionRequest request) {
            return base.Channel.ExecRemoteFunctionAsync(request);
        }
    }
}
