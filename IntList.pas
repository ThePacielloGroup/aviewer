unit IntList;

interface

uses Classes, Types;

type
    TIntegerList = class( TList )   {* IntegerList *}
    private
        FUnSigned   : Boolean;
        function    Get(Index: Integer): Integer;
        procedure   Put(Index: Integer; const Value: Integer);
    public
        function    Add(Num: Integer): Integer;
        function    First: Integer;
        function    IndexOf(Num: Integer): Integer;
        procedure   Insert(Index, Num: Integer);
        function    Last: Integer;
        function    Remove(Num: Integer): Integer;
        procedure   Pack(RemoveNum: Integer);
        procedure   Sort;

        property    Items[Index: Integer]: Integer read Get write Put; default;
        property    Nums[Index: Integer]: Integer read Get write Put;
        property    UnSigned: Boolean read FUnSigned write FUnSigned;
    end;

implementation

function SortBySmall(Item1, Item2: Pointer): Integer;
begin
    if Integer( Item1 )>Integer( Item2 ) then Result := 1
        else if Integer( Item1 )<Integer( Item2 ) then Result := -1
            else Result := 0;
end;

function SortByUnsignSmall(Item1, Item2: Pointer): Integer;
begin
    if Cardinal( Item1 )>Cardinal( Item2 ) then Result := 1
        else if Cardinal( Item1 )<Cardinal( Item2 ) then Result := -1
            else Result := 0;
end;

function TIntegerList.Get(Index: Integer): Integer;
begin
    Result := Integer( inherited Get( Index ) );
end;

procedure TIntegerList.Put(Index: Integer; const Value: Integer);
begin
    inherited Put( Index, Pointer( Value ) );
end;

function TIntegerList.Add(Num: Integer): Integer;
begin
    Result := inherited Add( Pointer( Num ) );
end;

function TIntegerList.First: Integer;
begin
    Result := Get( 0 );
end;

function TIntegerList.IndexOf(Num: Integer): Integer;
begin
    Result := inherited IndexOf( Pointer( Num ) );
end;

procedure TIntegerList.Insert(Index, Num: Integer);
begin
    inherited Insert( Index, Pointer( Num ) );
end;

function TIntegerList.Last: Integer;
begin
    Result := Get( Count-1 );
end;

function TIntegerList.Remove(Num: Integer): Integer;
begin
    Result := inherited Remove( Pointer( Num ) );
end;

procedure TIntegerList.Pack(RemoveNum: Integer);
var
    idx : Integer;
begin
    idx := 0;
    while idx<Count do begin
        if Get( idx )=RemoveNum then Delete( idx )
            else inc( idx );
    end;
end;

procedure TIntegerList.Sort;
begin
    if UnSigned then inherited Sort( SortByUnsignSmall )
        else inherited Sort( SortBySmall );
end;

end.
