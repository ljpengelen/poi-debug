package nl.cofx.poi.ticket;

import lombok.Builder;
import lombok.Value;

import java.time.LocalDate;
import java.util.Map;

@Builder
@Value
public class AssetPlanning {

    String assetId;
    String assetName;
    Map<LocalDate, String> statePerWeek;
}
