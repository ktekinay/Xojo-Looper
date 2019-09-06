#tag Class
Protected Class Looper_MTC
	#tag Method, Flags = &h0
		Sub Constructor()
		  MyTimer = new Timer
		  MyTimer.Mode = Timer.ModeOff
		  MyTimer.Period = 5
		  AddHandler MyTimer.Action, WeakAddressOf MyTimer_Action
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Destructor()
		  if MyTimer isa object then
		    Stop
		    RemoveHandler MyTimer.Action, WeakAddressOf MyTimer_Action
		    MyTimer = nil
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub MyTimer_Action(sender As Timer)
		  dim targetTicks as integer = Ticks + ActiveTicks
		  
		  while RaiseEvent MoveNext()
		    RaiseEvent HandleIteration
		    
		    if Ticks > targetTicks then
		      RaiseEvent Paused
		      sender.Period = PauseMilliseconds
		      sender.Mode = Timer.ModeSingle
		      return
		    end if
		  wend
		  
		  mIsRunning = false
		  RaiseEvent Finished
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 537461727420746865207461736B2E20496620697420616C72656164792072756E6E696E672C20746869732077696C6C20726169736520616E20556E737570706F727465644F7065726174696F6E457863657074696F6E2E
		Sub Start()
		  If mIsRunning then
		    raise new UnsupportedOperationException
		  end if
		  
		  mIsRunning = true
		  RaiseEvent Reset
		  MyTimer_Action MyTimer
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53746F7020612072756E6E696E67207461736B2E204974206973206E6F7420616E206572726F7220746F2063616C6C2074686973206576656E207768656E206E6F207461736B2069732072756E6E696E672E
		Sub Stop()
		  if MyTimer isa object then
		    MyTimer.Mode = Timer.ModeOff
		    mIsRunning = false
		  end if
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0, Description = 546865207461736B20697320636F6D706C6574652E
		Event Finished()
	#tag EndHook

	#tag Hook, Flags = &h0, Description = 50726F6365737320612073696E676C6520697465726174696F6E207573696E6720796F7572206C6F6F7020696E6465782C206574632E
		Event HandleIteration()
	#tag EndHook

	#tag Hook, Flags = &h0, Description = 496E6372656D656E7420746865206C6F6F7020696E6465782C206574632E20616E642072657475726E20547275652E20496620646F6E652C2072657475726E2046616C73652E2057696C6C2062652063616C6C6564206265666F726520746865207374617274206F66206561636820697465726174696F6E2C20696E636C7564696E67207468652066697273742E
		Event MoveNext() As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0, Description = 546865207461736B206861732070617573656420697465726174696F6E7320746F20616C6C6F77206F746865722070726F63657373696E672E2049742077696C6C20726573756D652061667465722050617573654D696C6C697365636F6E6473206D696C6C697365636F6E64732E
		Event Paused()
	#tag EndHook

	#tag Hook, Flags = &h0, Description = 546865207461736B2069732061626F757420746F2073746172742E20526573657420616E79206C6F6F7020636F756E746572732C206574632E2C2075736564207768696C652070726F63657373696E67206561636820697465726174696F6E206F6620746865206C6F6F702E
		Event Reset()
	#tag EndHook


	#tag Property, Flags = &h0, Description = 546865206D696E696E756D205469636B7320746F2072756E20697465726174696F6E73206265747765656E207061757365732E204174206C65617374206F6E6520697465726174696F6E2069732067756172616E746565642E
		ActiveTicks As Integer = 2
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mIsRunning
			End Get
		#tag EndGetter
		IsRunning As Boolean
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private mIsRunning As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private MyTimer As Timer
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 546865206D696E696D756D206D696C6C697365636F6E647320746F207061757365206265666F726520726573756D696E6720746865207461736B2E
		PauseMilliseconds As Integer = 10
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ActiveTicks"
			Visible=true
			Group="Behavior"
			InitialValue="3"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsRunning"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="PauseMilliseconds"
			Visible=true
			Group="Behavior"
			InitialValue="10"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
